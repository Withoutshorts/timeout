<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->

<%

func = request("func")

'*** Opdaterer job liste ****
if func = "ajaxupdate" then

Function AjaxFormatDates(datestring,interval)
if InStr(datestring,"/") then
splitslashdate = split(datestring,"/")
datestring = splitslashdate(2)&"-"&splitslashdate(1)&"-"&splitslashdate(0)
end if
datestring = DateAdd("d",interval,datestring)
datestringsplit = split(datestring,"-")
AjaxFormatDates = datestringsplit(2)&"/"&datestringsplit(1)&"/"&datestringsplit(0)
End Function

strAjaxStartdate = Request.Form("startdate")
strAjaxEnddate = Request.Form("enddate")
strAjaxCnr = Request.Form("cnr")
strAjaxJid = Request.Form("jid")
strAjaxDateDiff = Request.Form("DateDiff")
strAjaxId = Request.Form("id")
strAjaxCalcStartdate = AjaxFormatDates(strAjaxStartdate, strAjaxDateDiff)
strAjaxCalcEnddate = AjaxFormatDates(strAjaxEnddate, strAjaxDateDiff)
returnstring = "cnr= " & strAjaxCnr & " jid=" & strAjaxJid & " datediff= " & strAjaxDateDiff & " id=" & strAjaxId & " startdate =" & strAjaxStartdate & " enddate =" &strAjaxEnddate&" formattedstartdate ="'&AjaxFormatDates(strAjaxStartdate)
returnstring = returnstring & " added datediffstartdate = " & strAjaxCalcStartdate &" added datediffenddate = " & strAjaxCalcEnddate
response.Write returnstring

if strAjaxJid <> "undefined" then

'ustdato = replace(ustaar(u) & "/" & ustmd(u) & "/" & ustdag(u), ",", "")

	strSQLajaxjobupd = "UPDATE job SET jobstartdato = '"& strAjaxCalcStartdate &"', "_
	&" jobslutdato = '"& strAjaxCalcEnddate &"' WHERE id = " & strAjaxJid
	oConn.execute(strSQLajaxjobupd)
	
	strSQL2 = "SELECT aktstartdato, aktslutdato, id FROM aktiviteter WHERE job = " & strAjaxJid

	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF

	strSQLajaxaktupd = "UPDATE aktiviteter SET aktstartdato = '" & AjaxFormatDates(oRec2("aktstartdato"), strAjaxDateDiff) & "', "_
	&" aktslutdato = '" & AjaxFormatDates(oRec2("aktslutdato"), strAjaxDateDiff) & "' WHERE id = " & oRec2("id")
	oConn.execute(strSQLajaxaktupd)
	
	oRec2.movenext
	wend
	oRec2.close
	
	response.Write "updated job and activities"

elseif strAjaxId <> "undefined" then

	strSQLajaxaktupd = "UPDATE aktiviteter SET aktstartdato = '" & strAjaxCalcStartdate & "', "_
	&" aktslutdato = '" & strAjaxCalcEnddate & "' WHERE id = " & strAjaxId
	oConn.execute(strSQLajaxaktupd)

	response.Write "       start - " & strAjaxCalcStartdate & "   end -" & strAjaxCalcEnddate
	response.Write "updated job"

end if

Response.End
end if

if func = "opdaterjobliste" then


ujid = split(request("FM_jobid"), ",")

ustaar = split(request("FM_start_aar"))
ustmd = split(request("FM_start_mrd"))
ustdag = split(request("FM_start_dag"))

uslaar = split(request("FM_slut_aar"))
uslmd = split(request("FM_slut_mrd"))
usldag = split(request("FM_slut_dag"))

uudgifter = split(request("FM_udgifterpajob"), ",")
urest = split(request("FM_restestimat"), ",")
urisiko = split(request("FM_risiko"), ",")
ustatus = split(request("FM_jobstatus"), ",")


for u = 0 to UBOUND(ujid)
	'Response.write uudgifter(u) & "<br>"
	
	ustdato = replace(ustaar(u) & "/" & ustmd(u) & "/" & ustdag(u), ",", "")
	usldato = replace(uslaar(u) & "/" & uslmd(u) & "/" & usldag(u), ",", "")
	
	strSQLjobupd = "UPDATE job SET jobstartdato = '"& ustdato &"', "_
	&" jobslutdato = '"& usldato &"', "_
	&" udgifter = "& uudgifter(u) &", "_
	&" restestimat = "& urest(u) &", risiko = "& urisiko(u) &", "_
	&" jobstatus = "& ustatus(u) &" WHERE id = " & ujid(u)
	oConn.execute(strSQLjobupd)
	
	'Response.write strSQLjobupd & "<br>"
	
next



end if


thisfile = "webblik_joblisten2.asp"

if len(request("FM_parent")) <> 0 then
parent = request("FM_parent")
else
parent = 0
end if

if len(request("oldone")) <> 0 then
oldones = request("oldone")
else
oldones = 0
end if

if len(request("lastid")) <> 0 then
lastId = request("lastid")
else
lastId = 0
end if

print = request("print")






if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	%>
	
	

	<%if print <> "j" then 
	
	dtop = "132"
	dleft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.jobliste.js" type="text/javascript"></script>

	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(2)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call webbliktopmenu()
	%>
	</div>
	<%else %>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	dtop = "20"
	dleft = "20" 
	
	end if %>
	
	
	
	<%
	'*******************************************************
	'**** Igangværende Job *********************************
	'*******************************************************
	%>
	
	
	<div id="joblisten" style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px; visibility:visible; display:;">
	<h4>Timeout idag - Igangværende Job </h4>
	

	
	
	<%call filterheader(0,0,800,pTxt)%>
	
	
	<table width=100% cellpadding=0 cellspacing=0 border=0>
	<form method="post" action="webblik_joblisten2.asp?FM_usedatokri=1">
	<tr>
	<td valign=top>
	
	
	<% 
    if len(request("FM_kunde")) <> 0 then
			
			if request("FM_kunde") = 0 then
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			else
			valgtKunde = request("FM_kunde")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			end if
			
	else
		
			if len(request.cookies("webblik")("kon")) <> 0 AND request.cookies("webblik")("kon") <> 0 then
			valgtKunde = request.cookies("webblik")("kon")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			else
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			end if
			
	end if
	
	
	response.Cookies("webblik")("kon") = valgtKunde
    
    %>
    <b>Kontakter:</b>
    
    <%if print <> "j" then %>
    <br> <select name="FM_kunde" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
		<%
		end if
		
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
				
				if print <> "j" then%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		</select>
		<%end if %>
		
		
    <%if print = "j" then %>
    &nbsp;<%=vlgtKunde %><br />
    <%end if %>
    
    <br><br />
	
	<b>Start eller Slutdato?</b><br />
    <%
    if len(request("st_sl_dato")) <> 0 then
    usedatoKri = request("st_sl_dato")  
    else
        if len(request.Cookies("webblik")("datokri")) <> 0 then
        usedatoKri = request.Cookies("webblik")("datokri")
        else
        usedatoKri = 1
        end if  
    end if
    
    response.Cookies("webblik")("datokri") = usedatoKri
    
    
    st_sl_dato_forventet = 0
    if usedatoKri <> 0 then
    st_sl_DatoChk0 = ""
    st_sl_DatoChk1 = "CHECKED"
    stsldatoSQLKri = "jobslutdato"
            
            
            if len(request("st_sl_dato_forventet")) <> 0 AND request("st_sl_dato_forventet") <> "0" then
            stForDatoCHK = "CHECKED" 
            stsldatoSQLKri = "forventetslut"
            st_sl_dato_forventet = 1
            else
                
                
                    if len(request.Cookies("webblik")("forventet")) <> 0 then
                    st_sl_dato_forventet = request.Cookies("webblik")("forventet")
                            
                            if st_sl_dato_forventet = 1 then 
                            stForDatoCHK = "CHECKED" 
                            stsldatoSQLKri = "forventetslut"
                            st_sl_dato_forventet = 1
                            else
                            stForDatoCHK = ""
                            stsldatoSQLKri = stsldatoSQLKri
                            st_sl_dato_forventet = 0
                            end if
                            
                    else
                    stForDatoCHK = ""
                    stsldatoSQLKri = stsldatoSQLKri
                    st_sl_dato_forventet = 0
                    end if  
                    
            end if
            
            response.Cookies("webblik")("forventet") = st_sl_dato_forventet
            
    else
    st_sl_DatoChk0 = "CHECKED"
    st_sl_DatoChk1 = ""
    stsldatoSQLKri = "jobstartdato"
    stForDatoCHK = ""
    end if 
    
   
     
    %>
    
    <%if print <> "j" then %>
    <input id="st_sl_dato" name="st_sl_dato" type="radio" value="0" <%=st_sl_DatoChk0 %> /> Vis kun dem der har en <b>startdato</b> i det valgte datointerval<br />
	<input id="st_sl_dato" name="st_sl_dato" type="radio" value="1" <%=st_sl_DatoChk1 %>/>  Vis kun dem der har en <b>slutdato</b> i det valgte datointerval<br />
	<!--<input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="checkbox" value="1" <%=stForDatoCHK%> />&nbsp;Sorter efter <u>forventet</u> slutdato.<br />-->
	<%else 
	    
	    if st_sl_DatoChk0 = "CHECKED" then
	    Response.Write "- Startdato"
	    else
	    Response.Write "- Slutdato"
	    end if
	    
	end if %>
    
    
    <%if level = 1 then
        
        if len(request("FM_start_aar")) <> 0 then
        
        if len(request("FM_ignorer_pg")) <> 0 then
        IgnrProjGrp = 1
        else
        IgnrProjGrp = 0
        end if
        
		if cint(IgnrProjGrp) = 1 then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if
		
		else
		    
		    if request.cookies("erp")("prjgrp") <> "" then
		    IgnrProjGrp = request.cookies("erp")("prjgrp")
		        if cint(IgnrProjGrp) = 1 then
		        chkThis = "CHECKED"
		        else
		        chkThis = ""
		        end if
		     else
		     IgnrProjGrp = 0
		     chkThis = ""
		     end if
		
		end if
		
		response.cookies("erp")("prjgrp") = IgnrProjGrp
		
		%>
		<br><b>Projektgrupper:</b><br /><input type="checkbox" name="FM_ignorer_pg" value="1" <%=chkThis%>> Vis job uanset tilknytning til dine projektgrupper.
    <%end if%>
    <br /><br />
    <b>Søg på jobnr eller jobnavn:</b>&nbsp;<br />
		<input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=request("jobnr_sog")%>" style="width:258px;">
		<br />
    
	
	</td><td valign=top><b>Grafisk datomidtpunkt:</b><br />
	
	    <%if print = "j" then
	    dontshowDD = "1"
	    end if%>
	   
		<!--#include file="inc/weekselector_s.asp"-->
		
		<%if print = "j" then%>
		<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>
	    <%end if %>
	    
		
		
		<%
		if len(request("FM_valgrisiko")) <> 0 then
		risk = request("FM_valgrisiko")
		        if risk <> 0 then
		        riskSQlKri = " risiko = " & risk &" AND "
		        else
		        riskSQlKri = " risiko <> 4 AND "
		        end if
		else
		    if request.cookies("pga")("prioitet") <> "" then
		    risk = request.cookies("pga")("prioitet")
		        if risk <> 0 then
		        riskSQlKri = " risiko = " & risk &" AND "
		        else
		        riskSQlKri = " risiko <> 4 AND "
		        end if
		    else
		    risk = 0
		    riskSQlKri = " risiko <> 4 AND " 
		    end if
		    
		    
		end if 
		
		response.cookies("pga")("prioitet") = risk
		
		
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		
		prioTxt0 = "Alle (undt. Skjulte)" 
		prioTxt1 = "Lav" 
		prioTxt2 = "Mellem" 
		prioTxt3 = "Høj"
		prioTxt4 = "Skjulte" 
		
		select case cint(risk)
		case 1
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		vlgPrioTxt = prioTxt1
		
		case 2
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		stCHK3 = ""
		stCHK4 = ""
		vlgPrioTxt = prioTxt2
		
		case 3
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = "SELECTED"
		stCHK4 = ""
		vlgPrioTxt = prioTxt3
		
		case 4
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = "SELECTED"
		vlgPrioTxt = prioTxt4
		
		case else
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		vlgPrioTxt = prioTxt0
		
		end select
		
		'"& riskSQlKri &"
		%>
		
		
		<br /><br />
		<b>Prioitet:</b><br />
		
		<%if print <> "j" then %>
		
		<select name="FM_valgrisiko" id="FM_valgrisiko" style="font-size:9px; font-family:arial; width:190px;">
		<option value="0" <%=stCHK0%>><%=prioTxt0 %></option>
		<option value="1" <%=stCHK1%>><%=prioTxt1 %></option>
		<option value="2" <%=stCHK2%>><%=prioTxt2 %></option>
		<option value="3" <%=stCHK3%>><%=prioTxt3 %></option>
		<option value="4" <%=stCHK4%>><%=prioTxt4 %></option>
		</select>
		
		<%else %>
		
		- <%=vlgPrioTxt %>
		
		<%end if %>
		
		
		<br /><br />
		<b>Status:</b><br />
		- Aktive job<br />
		
		
		
		
		<%
        		
	   
	    if len(request("FM_start_aar")) <> 0 then
	        
	        if len(request("FM_status1")) <> 0 then
	        stat1 = 1
	        stCHK1 = "CHECKED"
	        else
	        stat1 = 0
	        stCHK1 = ""
	        end if
	        
	        if len(request("FM_status2")) <> 0 then
	        stat2 = 1
	        stCHK2 = "CHECKED"
	        else
	        stat2 = 0
	        stCHK2 = ""
	        end if
	        
	   
	    else
	        
	        if request.cookies("pga")("status1") <> "" then
	                
	                stat1 = request.cookies("pga")("status1")
	                if cint(stat1) = 1 then
	                stCHK1 = "CHECKED"
	                else
	                stCHK1 = ""
	                end if
	                
	        else
	        stat1 = 0
	        stCHK1 = ""
	        end if
	        
	        if request.cookies("pga")("status2") <> "" then
	                
	                stat2 = request.cookies("pga")("status2")
	                if cint(stat2) = 1 then
	                stCHK2 = "CHECKED"
	                else
	                stCHK2 = ""
	                end if
	                
	       
	        else
	        stat2 = 0
	        stCHK2 = ""
	        end if
	        
	    end if
	    
	    response.cookies("pga")("status1") = stat1
	    response.cookies("pga")("status2") = stat2
	    
	    if print <> "j" then%>
	    <input type="checkbox" name="FM_status1" value="1" <%=stCHK1%>>Passive job<br />
	     <input type="checkbox" name="FM_status2" value="2" <%=stCHK2%>>Lukkede job
	    <%else %>
	    
	        <%if stat1 = "1" then%>
	        - Passive job<br />
	        <%end if %>
	        
	         <%if stat2 = "1" then%>
	        - Lukkede job<br />
	        <%end if %>
	    
	    <%end if %>
	    
	    <%
	    response.cookies("pga").expires = date + 32
	    %>
		
		
    </td>
	<td valign=bottom>
	<%if print <> "j" then%>
	<input type="submit" value="Vis joblisten >>">
	<%end if%>
	</td></tr>
	</form>
	</table>
	
	<!-- filter div -->
	</td></tr></table>
	</div>
	
	<%if print <> "j" then %>
	            <br /><br /><table cellpadding=0 cellspacing=0 border=0><tr><td bgcolor="#ffffff" style="padding:0px 0px 0px 5px; border:1px #8cAAe6 solid; border-right:0px;">
	            <a href='jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik' class='vmenu' alt="Opret nyt job" target="_top">Opret nyt job</a>
                    </td><td bgcolor="#ffffff" style="padding:4px 5px 0px 5px; border:1px #8cAAe6 solid; border-left:0px;">
                    <a href='jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik' class='vmenu' alt="Opret nyt job" target="_top"><img src="../ill/add2.png" border="0" /></a>
                    </td></tr></table>
	            <%end if %>
	
	<br />
	<br />
	<table cellspacing=0 cellpadding=0 border=0 width=95%>
	<form method="post" action="webblik_joblisten2.asp?func=opdaterjobliste">
    
    <!--
    <input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="hidden" value="<%=st_sl_dato_forventet%>"/>
	-->
	
	<% 
	'Response.Write "strAAr = " & strAar
	sqlDatostartsplit = split(DateAdd("m",-6,strDag&"/"&strMrd&"/"&strAar),"-")
	sqlDatoslutsplit = split(DateAdd("m",6,strDag&"/"&strMrd&"/"&strAar),"-")
	sqlDatostart = sqlDatostartsplit(2) & "/" & sqlDatostartsplit(1) & "/" & sqlDatostartsplit(0)  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut =  sqlDatoslutsplit(2) & "/" & sqlDatoslutsplit(1) & "/" & sqlDatoslutsplit(0)'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	%>
	
    <%if print <> "j" then %>
	<tr>
	<td colspan=16 align=right style="padding:20px 10px 10px 0px;"><input type="submit" value="Opdater liste"></td>
    </tr>
	<%end if %>

<div id="JobGraph">
<ul id="JobGraphHeader">
<li class="infoField" style="width:200px;">
Kontakt (Knr.)
Jobnavn (Jobnr.)
</li>
<li style="border-right:0;">
</li>
</ul>
<div id="JobList">

</div>
	<center><img class="loadingIcon" src="../inc/jquery/images/ajax-loader.gif" alt="henter data"/></center>
</div>

</div>
	<%
	
	if request("FM_ignorer_pg") <> "1" then
		call hentbgrppamedarb(session("mid"))
	else
		strPgrpSQLkri = ""
	end if
	
	'*** Er der søgt på en specifikt jobnr/jobnavn ***
		jobnr_sog = request("jobnr_sog")
		if len(jobnr_sog) <> 0 AND jobnr_sog <> "-1" then
			jobnr_sog = jobnr_sog
			jobKri = " (j.jobnr = '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%') AND "
			show_jobnr_sog = jobnr_sog
			
		else
			jobnr_sog = "-1"
			jobKri = ""
			show_jobnr_sog = ""
		end if
	
	if cint(stat1) = 1 then
	statKri = "jobstatus = 1 OR jobstatus = 2"
	else
	statKri = "jobstatus = 1"
	end if
	
	if cint(stat2) = 1 then
	statKri = statKri & " OR jobstatus = 0 "
	else
	statKri = statKri
	end if
	
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobstatus, "_
	&" jobans2, kkundenr, jobslutdato, jobstartdato, j.beskrivelse, jobans1, jobans2, m.mnavn, m2.mnavn AS mnavn2,"_
	&" ikkebudgettimer, budgettimer, jobtpris, sum(r.timer) AS restimer, risiko, udgifter, forventetslut, restestimat, jobstatus "_
	&" FROM job j LEFT JOIN kunder ON (kid = jobknr) "_
	&" LEFT JOIN medarbejdere m ON (m.mid = jobans1)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
	&" LEFT JOIN ressourcer_md r ON (r.jobid = j.id)"_
	&" WHERE "& riskSQlKri &" fakturerbart = 1 AND ("& statKri &") AND " & jobKri & stsldatoSQLKri &" BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' "_
	& strPgrpSQLkri &" AND "& sqlKundeKri &""_
	&" GROUP BY j.id ORDER BY "& stsldatoSQLKri &", kkundenavn" 
	
	
	'Response.write strSQL
	'Response.flush
	
	c = 0
	budgetIalt = 0
	budgettimerIalt = 0
	restimerIalt = 0
	lastMonth = 0
	tilfakturering = 0
	udgifterTot = 0
	timerforbrugtIalt = 0
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	
	faktureret = 0
	timerTildelt = 0
	timerforbrugt = 0
	totalforbrugt = 0
	totalBalance = 0
	jobbudget = 0	
	
	'** Rettigheder til at redigere **
	if level = 1 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			end if
	end if


		timerTildelt = oRec("ikkebudgettimer") + oRec("budgettimer")
		if len(timerTildelt) <> 0 then
		timerTildelt = timerTildelt
		else
		timerTildelt = 0
		end if
		
		Response.write "<b>"& formatnumber(timerTildelt, 2) & "</b>"

		'*** Timer forbrugt ***
		strSQL2 = "SELECT sum(t.timer) AS timerforbrugt FROM timer t WHERE t.tjobnr = "& oRec("jobnr") &" AND t.tfaktim <> 5 GROUP BY t.tjobnr"
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		timerforbrugt = oRec2("timerforbrugt")
		end if
		oRec2.close
		
		if len(timerforbrugt) <> 0 then
		timerforbrugt = timerforbrugt
		else
		timerforbrugt = 0
		end if
		

		
		'*** Faktureret **'
		faktureret = 0
		timerFak = 0
		
		'*** Faktureret ***'
		strSQL2 = "SELECT sum(beloeb * (kurs/100)) AS faktureret, sum(timer) AS timer FROM fakturaer WHERE jobid = "& oRec("id") &" AND faktype = 0 GROUP BY jobid"
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		faktureret = oRec2("faktureret")
		timerFak = oRec2("timer")
		end if
		oRec2.close
		
		'*** Kredit ***'
		strSQL2 = "SELECT sum(beloeb * (kurs/100)) AS faktureret, sum(timer) AS timer FROM fakturaer WHERE jobid = "& oRec("id") &" AND faktype = 1 GROUP BY jobid"
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		faktureretKre = oRec2("faktureret")
		timerFakKre = oRec2("timer")
		end if
		oRec2.close
		
	
	%>
	<script type="text/javascript">
	ArrActivities = new Array();
	<%
	
	strSQL2 = "SELECT a.navn, a.id AS aid, a.budgettimer, a.fakturerbar, sum(t.timer) AS timer, t.tdato, t.timerkom, t.offentlig, "_
	&" t.sttid, t.sltid, "_
	&" a.incidentid, a.aktstartdato, a.aktslutdato, a.beskrivelse "_
	&" FROM aktiviteter a "_
	&" LEFT JOIN timer t ON (t.taktivitetid = a.id)"_
	&" WHERE a.job = " & oRec("id") & " GROUP BY a.id"
	
	ActivityNr = 0
	
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF
	
	
	//make realisedoutput no matter what
	if IsNull(oRec2("timer")) then
	Arealiserettimer = 0
	else
	Arealiserettimer = oRec2("timer")
	end if
	
	if IsNull(oRec2("budgettimer")) or oRec2("budgettimer") = 0 then
	//lad ikke budgettimer være lavere end realiseret timer
	if Arealiserettimer = 0 then
	Abudgettimer = 0
	else
	Abudgettimer = Arealiserettimer
	end if
	else
	Abudgettimer = oRec2("budgettimer")
	end if
	
	%>
	ArrActivities[<%=ActivityNr%>] = { "name" : <%="'"&JSFormat(oRec2("navn"))&"'"%>, "precalc" : <%="'"&JSFormat(Abudgettimer)&"'"%>, "realised" : <%="'"&JSFormat(Arealiserettimer)&"'"%>, "startdate" : <%="'"&oRec2("aktstartdato")&"'"%>, "enddate" : <%="'"&oRec2("aktslutdato")&"'"%>, "id" : <%="'"&oRec2("aid")&"'"%> };
	<%		
	ActivityNR = ActivityNR + 1
	oRec2.movenext
	wend
	oRec2.close
	
	
	 %>
																																		//AddJob(ArrActivities, customer, name, id, precalc,																																				realised, resource, startdate, enddate)
		AddJobList[<%=c%>] = {'ArrActivities' : ArrActivities, <%="'customer' : '"&JSFormat(oRec("kkundenavn")) &"' ,'name':'"&JSFormat(oRec("jobnavn"))&"','jnr': '"&oRec("jobnr")&"','jid': '"&oRec("id")&"','cnr': '"&oRec("kkundenr")&"','precalc': '"&JSFormat(timerTildelt)&"','realised': '"&JSFormat(timerforbrugt)&"','resource' : '"&JSFormat(oRec("restimer"))&"','startdate' : '"&oRec("jobstartdato")&"','enddate' : '"&oRec("jobslutdato")&"'"%>}
	</script>
	<%
	
	c = c + 1
	response.Flush
	oRec.movenext
	wend
	oRec.close 
		
		Function JSFormat(jsstring)
		
		IF LEN(TRIM(jsstring)) <> 0 then
        jsstring = replace(jsstring, ",",".")
        jsstring = replace(jsstring, "'","\'")
        JSFormat = jsstring
        else
        JSFormat = ""
        end if
	    
	    End Function
	%>
	
	</form>
	</table>
	
	</div>
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->