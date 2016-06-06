<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->

<%

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	func = request("func")
    thisfile = "webblik_tilfakturering.asp"
    print = request("print")
	
	
	if print <> "j" then
	
	dTop = "132"
	dLeft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(2)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call webbliktopmenu()
	%>
	</div>
	
	<%else 
	
	dTop = "20"
	dLeft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<%end if %>
	
	
	<div id="sindhold" style="position:absolute; left:<%=dLeft %>px; top:<%=dTop %>px; visibility:visible;">
	<%
	'**********************************************************
	'**************** Job til fakturering *********************
	'**********************************************************
	
	
	
	
	%>
	<h4>TimeOut idag - Job til fakturering</h4>
	<table width=60% cellspacing=0 cellpadding=0 border=0>
	<form action="webblik_tilfakturering.asp?FM_usedatokri=1" method="POST"><tr>
	<td style="border-left:1px limegreen solid; border-top:1px limegreen solid; border-right:2px forestgreen solid; border-bottom:2px forestgreen solid; padding:10px;" bgcolor="#ffffff" valign=top>
	
    <% 
    if len(request("FM_kunde")) <> 0 then
			
			if request("FM_kunde") = 0 then
			valgtKunde = 0
			sqlKundeKri = "t.tknr <> 0"
			else
			valgtKunde = request("FM_kunde")
			sqlKundeKri = "t.tknr = "& valgtKunde &""
			end if
			
	else
		
			if len(request.cookies("webblik")("kon")) <> 0 AND request.cookies("webblik")("kon") <> 0 then
			valgtKunde = request.cookies("webblik")("kon")
			sqlKundeKri = "t.tknr = "& valgtKunde &""
			else
			valgtKunde = 0
			sqlKundeKri = "t.tknr <> 0"
			end if
			
	end if
	
	
	response.Cookies("webblik")("kon") = valgtKunde
    
    %>
    <b>Kontakter:</b><br> <select name="FM_kunde" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
		<%
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(valgtKunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select><br><br />
	    
	     <%if level = 1 then
		if request("FM_ignorer_pg") = "1" then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if
		%>
		<br><b>Projektgrupper:</b><br /><input type="checkbox" name="FM_ignorer_pg" value="1" <%=chkThis%>> Vis job og timer til fakturering, uanset tilknytning til dine projektgrupper.<br /><br />
        <%end if%>
	    
	    
	    <%  
	    if request("FM_ignorer_pg") <> "1" then
		call hentbgrppamedarb(session("mid"))
	    else
		strPgrpSQLkri = ""
	    end if
	    %>
	    
	   <b>Periode:</b><br /> 
	    <!--#include file="inc/weekselector_s.asp"-->&nbsp;<br />
	   <% if print <> "j" then%>
       <br /> <input id="Submit1" type="submit" value="  Vis timer til fak.  " />
	<%else %>
	&nbsp;<%=chkNameValue %>
	<%end if %>
	<br /><br />
	Hvis der findes en faktura på de pågældende job i det valgte dato-interval medreges kun timer siden sidste faktura.
	    </td></tr></form></table>
	<%
	
	 
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	%>
	<br><b>Periode:</b>&nbsp;
	<%=formatdatetime(strDag&"/"& strMrd &"/"& strAar, 1) & " - " & formatdatetime(strDag_slut &"/"& strMrd_slut &"/"& strAar_slut, 1)%>
	
	<%if print <> "j" then%>
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="webblik_tilfakturering.asp?print=j" class=vmenu target=blank>Print venlig version</a>
	<%end if%>
	
	<table cellspacing=0 cellpadding=0 border=0 width=95%>
	<tr>
		 <td colspan=2 align=right style="padding:2px;">&nbsp;</td>
	</tr>
	<%
	
	strSQLFAK = "SELECT f.fakdato FROM fakturaer WHERE "
	
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobans2, "_
	&" kkundenr, jobslutdato, jobstartdato, "_
	&" j.beskrivelse, jobans1, mnavn, tjobnr, sum(t.timer) AS timer, "_
	&" kid, j.serviceaft, s.navn AS aftalenavn FROM timer t, job j "_
	&" LEFT JOIN kunder ON (kid = jobknr)"_
	&" LEFT JOIN medarbejdere m ON (mid = jobans1)"_
	&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft)"_
	&" WHERE tfaktim = 1 AND (jobnr = tjobnr "& strPgrpSQLkri &") AND "& sqlKundeKri &""_
	&" AND tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"_ 
	&" GROUP BY t.tjobnr ORDER BY kkundenavn, jobnavn, jobslutdato"  
	
	
	
	'Response.write strSQL
	'Response.flush
	
	
	c = 0
	jobIdstr = "#0#"
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	if oRec("serviceaft") <> 0 then
	bgthis = "#eff3ff"
	else
		'bg = right(c, 1)
		'select case bg
		'case 0,2,4,6,8
		bgthis = "#ffffff"
		'case else
		'bgthis = "WhiteSmoke"
		'end select
	end if
	
	
	'if len(oRec("faknr")) > 1 then
	
	'else
	
	
	if cint(instr(jobIdstr, "#"&oRec("tjobnr")&"#")) = 0 then
	jobIdstr = jobIdstr &",#"&oRec("tjobnr")&"#"
	
	if lastKnr <> oRec("kkundenr") then%>
	<tr>
		 <td height="20" colspan="2">&nbsp;</td>
	</tr>
	<tr bgcolor="#ffff99">
		<td colspan=2 style="padding:15px 0px 3px 6px;"><b><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)</b></td>
	</tr>
	<%end if %>
	
	<tr bgcolor="<%=bgthis%>">
		
		<td valign=top style="padding:5px 0px 3px 6px; border-top:1px #2c962d dashed;">
		
		
		<% if print <> "j" then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=webblik_tilfakturering" class=todo_mellem><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</a>
		<%else %>
		<%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)
		<%end if %>
		<br>
		<%if oRec("serviceaft") <> 0 then%>
		Aftale: <%=oRec("aftalenavn")%><br>
		<%end if%>
		<font color="green"><%=formatdatetime(oRec("jobstartdato"), 1)%></font> til <font color="red"><%=formatdatetime(oRec("jobslutdato"), 1)%></font> | <font color="#c4c4c4"><i><%=oRec("mnavn")%></i></font>
		</td>
		
		
		<td valign=top  style="padding:5px 0px 3px 6px; border-top:1px #2c962d dashed;">
		
		<%
		lastFak = "01/01/01"
		fakfindes = 0
		lastFaknr = 0
		
		strSQLFak = "SELECT f.fakdato, f.faknr FROM fakturaer f "_
		&" WHERE (f.jobid = "& oRec("id") &" AND "_
		&" f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') "_
		&" ORDER BY f.fakdato DESC"
		
		'Response.Write strSQLFak &"<br>" 
		
		oRec3.open strSQLFak, oConn, 3
        if not oRec3.EOF then
            fakfindes = 1
            lastFak = oRec3("fakdato")
            lastFaknr = oRec3("faknr")
        end if
        oRec3.close
                
                
                lastFakuseSQL = dateadd("d", 1, lastFak)
                lastFakuseSQL = year(lastFakuseSQL) &"/"& month(lastFakuseSQL) &"/"& day(lastFakuseSQL)
                sumTimerVedFak = 0
                if fakfindes = 1 then
                strSQLtimer = "SELECT sum(timer) AS sumtimerEfterFak FROM timer WHERE "_
                &" tfaktim = 1 AND tjobnr = "& oRec("jobnr") &""_ 
	            &" AND tdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY tjobnr"
	            
	           'Response.Write strSQLtimer &"<br>"
	            
	            oRec2.open strSQLtimer, oConn, 3
                while not oRec2.EOF
                
                sumTimerVedFak = oRec2("sumtimerEfterFak")
                
                oRec2.movenext
                wend
                oRec2.close
                
                end if
		
		
		 if print <> "j" then%>
		    
		    <%if oRec("serviceaft") = 0 then%>
		    <a href="stat_fak.asp?menu=stat_fak&FM_kunde=<%=oRec("kid")%>&shokselector=1&fakdenne=<%=oRec("jobnr")%>" class=todo_mellem>Opret ny faktura &nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		    <%else%>
		    <a href="fak_serviceaft_osigt.asp?menu=stat_fak&func=osigtall&id=<%=oRec("kid")%>" class=todo_mellem>Opret ny faktura &nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		    <%end if%>
		
		<%else%>
		Til Fakturering:
		<%end if%>
		
		
		<%if fakfindes = 1 then%>
		 <br>(Timer til fakturering: <b><%=sumTimerVedFak%></b>)
		 - Seneste faktura i valgt periode: <%=lastFaknr%>, <%=lastFak%>
		<%else %>
		 <br>(Timer til fakturering: <b><%=oRec("timer")%></b>)
		<%end if %>
		</td>
		
	</tr>
	
	<%
	lastKnr = oRec("kkundenr")
	c = c + 1
	
	end if
	
	
	'end if
	oRec.movenext
	wend
	oRec.close 
	
	
	if c = 0 then
	%>
	<tr><td style="padding:20px;">
	<b>Info:</b><br />
	Der blev ikke fundet nogen job med timer på i det valgte interval.<br />
	Vær opmærksom på at du kan ignorere dine tilhørsforhold til job via projektgrupper fra og prøve igen.</td></tr>
	<%
	end if
	%></table>
	<br /><br />
        &nbsp;
	

	</div>
	
	
	
	
	
	
	
	
	
	
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->