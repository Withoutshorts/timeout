	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<SCRIPT language=javascript src="inc/timereg_2006_func.js"></script>
	
	<%
	'*********** Sætter sidens global variable ***********************************'
	
	'* Hvis der vælges en anden bruger *'
	'* Den den der er logget på. *'
	 
	'*** Mid ***'
	if len(Request("FM_use_me")) <> 0 then
	usemrn = Request("FM_use_me")
	else
	usemrn = session("Mid")
	end if
	
	thisfile = "timereg"
	func = request("func")
	
	
	
	if func = "skiftmedarb" then
	%>
	<script>
	window.top.frames['a'].location.href = "timereg_akt_2006.asp?usemrn=<%=usemrn%>&showakt=1"
	</script>
	<%
	end if
	
	level = session("rettigheder")
	
	'***********************************
	'*** flytter job til guide (Skjuler job) ****
	moveJobtoguide = request("FM_flyttilguide")
	if len(moveJobtoguide) <> 0 then
		j = 0
		intuseJob = Split(moveJobtoguide, ", ")
		For j = 0 to Ubound(intuseJob)
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& usemrn &" AND jobid = "& intuseJob(j) &"")
		next
	end if
	
		
	'************************************************************************************************
	'*** Viser Timereg *******
	'************************************************************************************************
	%>
	
	<!--include file="inc/timereg_func_inc.asp"-->
	
<%

function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
end function

function SQLBless2(s)
		dim tmp2
		tmp2 = s
		tmp2 = replace(tmp2, ".", ",")
		SQLBless2 = tmp2
end function


%>

<script>
    $(document).ready(function() {

    $("#pagesetadv_but").click(function() {


        $("#pagesetadv").css("display", "");
        $("#pagesetadv").css("visibility", "visible");
        $("#pagesetadv").show(1000);
        

    });

    $("#pagesetadv_close").click(function() {


        $("#pagesetadv").hide("fast");


    });
    
    });

</script>

   <!---- Side indstillinger --->
	   <div id="pagesetadv" style="position:absolute; display:none; visibility:hidden; top:190px; left:20px; width:180px; height:155px; z-index:10000000000; background-color:#ffffff; border:1px #999999 solid; padding:3px;">
        
        <table cellspacing=0 cellpadding=2 border=0 bgcolor="#FFFFFF" width=100%>
         <form action="timereg_2006.asp?func=settings&FM_kontakt=<%=sogKontakter%>&FM_sog_job_navn_nr=<%=show_sogjobnavn_nr%>&FM_ignorer_projektgrupper=<%=ignorerProgrp%>" method="POST" name="timereg">
	   <tr bgcolor="#8caae6">
	   <td><img src="../ill/gear.png"/></td>
	   <td class=alt align=left><b><%=tsa_txt_307%></b></td>
	   <td align=right><a href="#" id="pagesetadv_close" class=red>[x]</a></td></tr>
        <tr><td colspan=3>
        
		<b><%=tsa_txt_308 %>:</b><br /> 
		<%
		t_skjuljobnr = 0
		if len(trim(request("FM_skjuljobnr"))) <> 0 then
		    
		    t_skjuljobnr = request("FM_skjuljobnr")
		else
		    if request.Cookies("timereg_2006")("t_skjuljobnr") <> "" then
		    t_skjuljobnr = request.Cookies("timereg_2006")("t_skjuljobnr")
		    else
		    t_skjuljobnr = 0
		    end if
	    end if
		
		    if cint(t_skjuljobnr) = 0 then
		    chk0 = "CHECKED"
		    chk1 = ""
		    else
		    chk1 = "CHECKED"
		    chk0 = ""
		    end if  
		
		response.Cookies("timereg_2006")("t_skjuljobnr") = t_skjuljobnr
		
		%>
		<input type="radio" name="FM_skjuljobnr" value="0" <%=chk0%>> <%=tsa_txt_314 %>
		<input type="radio" name="FM_skjuljobnr" value="1" <%=chk1%>> <%=tsa_txt_315 %>
		
        <br>
		
		
		<br /><b><%=tsa_txt_309 %>:</b><br /> 
		<%
		t_sorter = 0
		if len(trim(request("FM_sorter"))) <> 0 then
		    
		    t_sorter = request("FM_sorter")
		else
		    if request.Cookies("timereg_2006")("t_sorter") <> "" then
		    t_sorter = request.Cookies("timereg_2006")("t_sorter")
		    else
		    t_sorter = 0
		    end if
	    end if
		
		    select case cint(t_sorter)
		    case 1 
		    chk1 = "CHECKED"
		    chk0 = ""
		    chk2 = ""
		    case 2
		    chk0 = ""
		    chk1 = ""
		    chk2 = "CHECKED"
		    case else
		    chk0 = "CHECKED"
		    chk1 = ""
		    chk2 = ""
		    end select  
		
		response.Cookies("timereg_2006")("t_sorter") = t_sorter
		
		%>
		<input type="radio" name="FM_sorter" value="0" <%=chk0%>> <%=tsa_txt_310 %><br />
		<input type="radio" name="FM_sorter" value="1" <%=chk1%>> <%=tsa_txt_311 %><br />
		<input type="radio" name="FM_sorter" value="2" <%=chk2%>> <%=tsa_txt_316 %><br />
		
		
		
		
		<br /><b><%=tsa_txt_313 %>:</b><br /> 
		<%
		t_lickid = 0
		if len(trim(request("FM_lickid"))) <> 0 then
		    
		    t_lickid = request("FM_lickid")
		else
		    if request.Cookies("timereg_2006")("t_lickid") <> "" then
		    t_lickid = request.Cookies("timereg_2006")("t_lickid")
		    else
		    t_lickid = 0
		    end if
	    end if
		
		    if cint(t_lickid) = 0 then
		    chk0 = "CHECKED"
		    chk1 = ""
		    else
		    chk1 = "CHECKED"
		    chk0 = ""
		    end if  
		
		response.Cookies("timereg_2006")("t_lickid") = t_lickid
		
		%>
		<input type="radio" name="FM_lickid" value="0" <%=chk0%>> <%=tsa_txt_054 %><br />
		<input type="radio" name="FM_lickid" value="1" <%=chk1%>> <%=tsa_txt_055 %>
		
		
        <br>&nbsp;
		
	   </td>
	   </tr>
	   <tr>
	   <td colspan=3 align=right><input type="submit" value="<%=tsa_txt_086 %> <> " style="height:18px; font-family:arial; font-size:10px;"></td>
	   </tr>
	   	</form>
	   </table>
	   </div>


	<div id="sindhold" style="position:absolute; left:10; top:0; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	
<%

  

 
 
 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	            
 	            '*** Søge Scenarier ***
 	            '1) Jobnavn / nr
 	            '2) Kontakt
 	            '3) Kombination af ovenstående
 	            '4) Ignorer projektgruppe tilknytning.
 	            
 	             
 	            
				'*** Tilvalgte job i guiden dine aktive job ****
				varUseJob = "("
				
				
				if len(request("FM_sog_job_navn_nr")) <> 0 then
				sogjobnavn_nr = request("FM_sog_job_navn_nr")
				show_sogjobnavn_nr = sogjobnavn_nr
				varAktivejob = "(jobid = 0 OR "
				else
				sogjobnavn_nr = 0
				show_sogjobnavn_nr = ""
				varAktivejob = "("
				end if
				
				if len(request("FM_kontakt")) <> 0 then
				sogKontakter = request("FM_kontakt")
				varAktivejob = "(jobid = 0 OR "
				
				
				response.cookies("timereg_2006")("kontakt") = sogKontakter
	            response.cookies("timereg_2006").expires = date + 132
				
				else
				    
				        if request.Cookies("timereg_2006")("kontakt") <> "" AND request.Cookies("timereg_2006")("kontakt") <> 0 then
				        sogKontakter = request.Cookies("timereg_2006")("kontakt")
				        varAktivejob = "(jobid = 0 OR "
				        else
				        sogKontakter = 0
				        varAktivejob = "("
				        end if
				end if
				
				
				if len(request("FM_ignorer_projektgrupper")) <> 0 then
				ignorerProgrp = request("FM_ignorer_projektgrupper") '1
					if cint(ignorerProgrp) = 1 then
					selIgn = "CHECKED"
					else
					selIgn = ""
					end if
				else
				ignorerProgrp = 0
				selIgn = ""
				end if
				
				'*****************************************
				'**** Søger på job ***********************
				'*** Kontakter "overruler" søg på job **** 
				'*****************************************
				
				
				
				if len(trim(show_sogjobnavn_nr)) <> 0 then
				strJobSogKri = " AND (jobnr = '"& sogjobnavn_nr &"' OR jobnavn LIKE '"& sogjobnavn_nr &"%') "
				else
				strJobSogKri = ""
				end if
				
				if cint(sogKontakter) <> 0 then
				strKontaktKri = "jobknr = "& sogKontakter & " "
				else
				strKontaktKri = "jobknr <> 0 "
				end if
				
				
				 
					strSQL3 = "SELECT id, jobnr, jobnavn FROM job WHERE "& strKontaktKri &" "& strJobSogKri &" AND jobstatus = 1"
					'Response.write strSQL3
					'Response.Flush
					oRec3.open strSQL3, oConn, 3
					
					while not oRec3.EOF 
					varUseJobSog = varUseJobSog & " j.id = "& oRec3("id") & " OR "
					varAktivejob = varAktivejob & " jobid = "& oRec3("id") & " OR " 
					oRec3.movenext
					wend 
					
					oRec3.close 
					
					'sogKontakter = 0
				
				
				'********************************
				
				
				
				
				'****************************************
				'*** Skal projektgrupper ignoreres ?? ***
				'if ignorerProgrp = 0 then
					
					if varAktivejob = "(" then
					varAktivejob = ""
					else
					varAktivejob_len = len(varAktivejob)
					varAktivejob_left = varAktivejob_len - 3
					varAktivejob_use = left(varAktivejob, varAktivejob_left) 
					varAktivejob = varAktivejob_use & ") AND "
					end if
				
				'else
				
					'strPgrpSQLkri = ""
					'varAktivejob = ""
					
				'end if
				
				
				
				
				if ignorerProgrp = 0 then
				
				
						'******************************************************
						'**** Henter job fra Guiden Dine aktive job ***********
						strSQL4 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE "& varAktivejob &" medarb = "& usemrn &""
						oRec3.open strSQL4, oConn, 3
						
						while not oRec3.EOF 
						varUseJob = varUseJob & " j.id = "& oRec3("jobid") & " OR "
						oRec3.movenext
						wend 
						
						oRec3.close 
				
				else
				
				varUseJob = " ("& varUseJobSog	
						
				end if
				
				'Response.write varUseJob 
				'Response.flush
				
				
				if trim(varUseJob) = "(" then
				varUseJob = " j.id = 0 AND "
				else
				varUseJob_len = len(varUseJob)
				varUseJob_left = varUseJob_len - 3
				varUseJob_use = left(varUseJob, varUseJob_left) 
				varUseJob = varUseJob_use & ") AND "
				end if
				
				
				
				'**** Henter projektgrupper ***
				call hentbgrppamedarb(usemrn)
				
	
	'**************************************************************************************************
	
	
	call browsertype()
	if browstype_client = "mz" then
	dwdt = 185
	else
	dwdt = 195
	end if
	
	pTxt = replace(tsa_txt_115, "|", "&")
	
	call filterheader(0,0,dwdt,pTxt)
	
	
	
	'Response.Write "sogjobnavn_nr"& sogjobnavn_nr & "<br>"
	'Response.Write "sogKontakter" & sogKontakter
				
	
	%>
	
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<form action="timereg_2006.asp?FM_use_me=<%=usemrn%>" name="sogekri" method="POST">
	
	<tr>
		<td><b>A)&nbsp;</b><%=tsa_txt_073 %>:<br>
		<input type="text" name="FM_sog_job_navn_nr" id="FM_sog_job_navn_nr" value="<%=show_sogjobnavn_nr%>" style="font-family:arial; font-size:10px; width:140px;"> 
		<input type="checkbox" name="FM_ignorer_projektgrupper" id="FM_ignorer_projektgrupper" value="1" <%=selIgn%>><%=tsa_txt_074 %>
		</td>
	</tr>
	<tr>
	<%
	'**** Finder kunder til dd *****'
		strKundeKri = " AND ("
		strSQL3 = "SELECT j.id, medarb, jobid, kid, jobknr FROM timereg_usejob "_
		&"LEFT JOIN job j ON (j.id = jobid) "_
		&"LEFT JOIN kunder ON (kid = j.jobknr) "_
		&"WHERE medarb = "& usemrn &" AND kid <> '' ORDER BY kid"
		'response.write strSQL3
		'Response.flush
		oRec3.open strSQL3, oConn, 3
		
		while not oRec3.EOF 
		strKundeKri = strKundeKri & " kid = "& oRec3("kid") & " OR "
		oRec3.movenext
		wend 
		
		oRec3.close 
	%>
	<td style="padding-top:5px; padding-bottom:5px;"><b>B)</b>&nbsp;<%=tsa_txt_075 %>:
	<br>
	<select name="FM_kontakt" style="font-family:arial;font-size:10px; width:120px;"> <!-- onChange="renssog()" --> 
		<option value="0"><%=tsa_txt_076 %></option>
		<%
		if len(strKundeKri) <> 0 then
		strKundeKri = strKundeKri &" kid = 0)"
		else
		strKundeKri = strKundeKri &" AND (kid = 0)"
		end if
		
		ketypeKri = " ketype <> 'e'"
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(sogKontakter) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
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
		
		</select>&nbsp;&nbsp;<input type="submit" value="<%=tsa_txt_078 %>" style="width:50px; height:18px; font-family:arial; font-size:9px;">
			
	</td>
	
	</tr>
	</form>
	
	<form action="timereg_2006.asp?func=skiftmedarb&FM_kontakt=<%=sogKontakter%>&FM_sog_job_navn_nr=<%=show_sogjobnavn_nr%>&FM_ignorer_projektgrupper=<%=ignorerProgrp%>" name="selmed" method="POST">
	<%if level = 1 then
	SQLkriEkstern = ""
	%>
	<tr bgcolor="#ffffff">
	<td style="padding-top:5px; padding-bottom:5px; border-top:1px #003399 dashed; border-bottom:1px #003399 dashed;" class=lille><%=tsa_txt_077 %>
	<select name="FM_use_me" style="font-family: arial; font-size: 9px; width:90px;">
				<%
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat <> 2 AND mansat <> 3 "& SQLkriEkstern &" ORDER BY Mnavn"
					oRec.open strSQL, oConn, 3
					
					while not oRec.EOF 
					
					if cint(oRec("Mid")) = cint(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if%>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn")%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>
	
	&nbsp;<input type="submit" value="<%=tsa_txt_079 %>" style="height:18px; font-family:arial; font-size:9px;"></td>
	</tr>
	</form>
	<%end if%>
	<tr><td style="padding-top:5px;" colspan=2>
	<a href="javascript:popUp('stopur_2008.asp','1000','650','20','20')" class=rmenu><img src="../ill/stopur_32.png" alt="<%=tsa_txt_080 %>" border=0 /></a>
    &nbsp;&nbsp;
	<img src="../ill/ac0052-24.gif" width="24" height="24" alt="<%=tsa_txt_081 %>" border="0">&nbsp;<a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>','600','500','150','120');" target="_self" class=rmenu><%=tsa_txt_082 %></a>
	</td></tr>
	</table>
	</td></tr></table>
	</div>
	
	<%if level <= 3 OR level = 6 then %>
	<%
	
	oleftpx = 0
	otoppx = 20
	owdtpx = 110
	
	call opretNy("jobs.asp?menu=job&func=opret&id=0&int=1&rdir=treg", "Opret nyt job", otoppx, oleftpx, owdtpx) 
	%>
	<br />
	<%end if %>

	<br /><br />
	<b><%=replace(tsa_txt_083, "|", "&") %>:</b> &nbsp; <a href="#" id="pagesetadv_but" class="vmenu">+ <%=tsa_txt_302%></a>
	
	<div name="kontakterogjob" id="kontakterogjob" style="position:relative; left:0px; top:0px; display:; width:<%=dwdt%>px; visibility:visible; background-color:#ffffff; padding:3px; border:1px #8caae6 solid;">
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<form>
	<%
	
	'************** Job med adgang til (ved ignorer projektgrupper) ******************************************
	
	
	
	if ignorerProgrp = 1 then
	
	strJobmedadgang = "#0#"
	
		strSQL = "SELECT j.id AS id FROM job j "_
		&" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 "& strPgrpSQLkri & " ORDER BY j.id" 
		
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		strJobmedadgang = strJobmedadgang &",#"& oRec("id")&"#"
		
		oRec.movenext
		wend
		
		oRec.close
	
	end if
	
	
	if cint(t_lickid) <> 0 then
	'*** Licens indehaver **'
	call licKid()
	z = 2
	else
	z = 1
	end if
	
	lastKid = 0
	x = 0
	
	for s = 1 to z
	
	
	'************** Main SQL Call ******************************************
	strSQL = "SELECT j.id AS id, jobnr, jobnavn, jobslutdato, jobknr, kkundenavn, kkundenr, count(a.id) AS antalakt,"_
	&" j.budgettimer, j.ikkebudgettimer, "_
	&" kid, jobans1, jobans2 FROM job j, kunder k"_
	&" LEFT JOIN aktiviteter a ON (a.job = j.id)"
	
	if ignorerProgrp = 0 then
	strSQL = strSQL &" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 AND k.Kid = j.jobknr "& strPgrpSQLkri & "" 
	else
	strSQL = strSQL &" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 AND k.Kid = j.jobknr " 
	end if
	
	if cint(s) = 1 AND cint(t_lickid) <> 0 then
	strSQL = strSQL &" AND k.kid = "& licensindehaverKid  
	else
	    if cint(s) = 2 then
	    strSQL = strSQL &" AND k.kid <> "& licensindehaverKid  
	    else
	    strSQL = strSQL &""
	    end if 
	    
	    
	    '***+/- 2 år. kun på eksterne job, hvis vis interne job øverst er valgt ***'
	    if (t_sorter) = 2 then
	
	    dd = now
	    ddLow = dateadd("yyyy", -2, dd)
	    ddHigh = dateadd("yyyy", 2, dd)
    	
	    ddLowSQLformat = year(ddLow) & "/" & month(ddLow) & "/" & day(ddLow)
	    ddHighSQLformat = year(ddHigh) & "/" & month(ddHigh) & "/" & day(ddHigh)   
    	
	    strSQL = strSQL &" AND jobslutdato BETWEEN '"& ddLowSQLformat &"' AND '"& ddHighSQLformat &"'"
	    end if
	    
	end if
	
	
	
	select case cint(t_sorter)
	case 1
	strSQL = strSQL &" GROUP BY j.id, a.job ORDER BY j.jobnr DESC LIMIT 50"
	case 2
	strSQL = strSQL &" GROUP BY j.id, a.job ORDER BY j.jobslutdato DESC"
	case else
	strSQL = strSQL &" GROUP BY j.id, a.job ORDER BY k.kkundenavn, j.jobnavn, j.id, j.jobnavn"
    end select
	
	'Response.write strSQL
	'Response.flush
	'Response.write "<hr>"
	
	
	
	oRec.Open strSQL, oConn, 3
	
	xi = 0
	while not oRec.EOF 
	
	editok = 0
	
	if cint(t_sorter) = 1 then
		
		select case right(x,1)
		case 0,2,4,6,8
		bgThis = "#ffffff"
		case else
		bgThis = "#EFf3ff"
		end select
		
	else
	bgThis = "#ffffff"	
    end if
	
	
	
		'** rettigheder ** 
		if level <= 2 OR level = 6 then
		editok = 1
		else
				if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR _
				(cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) OR lto = "acc" then
				editok = 1
				end if
		end if
	
	'**lyserød streg efer interne jobs ***'	
	if cint(s) = 2 AND xi = 0 then%>
	<tr bgcolor="#ffdfdf">
		<td colspan=2 style="height:3px;"><img src="../ill/blank.gif" width="1" height="1" /></td>
	</tr>
	    
	     <%if cint(t_sorter) = 1 AND cint(t_lickid) = 1 then  %>
	     <tr bgcolor="#eff3ff">
		    <td colspan=2 style="padding:2px; border-bottom:1px silver solid;"><b><%=tsa_txt_311%> </b></td>
	    </tr>
	     <%end if %>
	    
	<%end if
	
	
	
	if (lastKid <> oRec("kid") AND cint(t_sorter) = 0) OR (cint(s) = 1 AND cint(t_lickid) <> 0 AND x = 0) then%>
	<tr bgcolor="#eff3ff">
		<td colspan=2 style="padding:2px; border-bottom:1px silver solid;"><b><%=oRec("kkundenavn")%> </b>(<%=oRec("kkundenr")%>)</td>
	</tr>
	<%end if
	
	if (cint(s) = 1 AND cint(t_lickid) = 0) OR cint(s) = 2 then
	
	if (month(lastJobslutdato) <> month(oRec("jobslutdato")) OR year(lastJobslutdato) <> year(oRec("jobslutdato"))) AND cint(t_sorter) = 2 then%>
	<tr bgcolor="#eff3ff">
		<td colspan=2 style="padding:2px; border-bottom:1px silver solid;"><b><%=monthname(month(oRec("jobslutdato"))) & " " & year(oRec("jobslutdato"))%></b></td>
	</tr>
	<%end if%>
	
	<%end if%>
	
	<tr bgcolor="<%=bgThis%>" id="tr_<%=oRec("id")%>">
		
		
		
		
		<%
		if (instr(strJobmedadgang, "#"& oRec("id") &"#") <> 0) OR ignorerProgrp = 0 then
		projektgrpOK = 1
		
			strprojektgrpOK = "&nbsp;"
		
		else
		projektgrpOK = 0
			if editok = 1 then
			strprojektgrpOK = "<a href=""javascript:popUp('tilknytprojektgrupper.asp?id="&oRec("id")&"&medid="&usemrn&"','600','500','150','120');"" target=""_self""; class=rmenu><IMG SRC='../ill/ac0021-16.gif' border='0' width='16' height='16' alt='Tilføj projektgrupper' /></a>"
			else
			
			strprojektgrpOK = "&nbsp;"
			end if
		end if
		
		
		
		%>
		
		
		<td style="padding:2px 2px 2px 2px; border-bottom:1px silver dashed;" class=lille>
		<%if projektgrpOK = 1 then%>
		<a href="timereg_akt_2006.asp?jobid=<%=oRec("id")%>&usemrn=<%=usemrn%>&showakt=1" target="a" onClick=markerjob(<%=oRec("id")%>); class=rmenu>
		<%=oRec("jobnavn")%>
		<%if cint(t_skjuljobnr) <> 1 then %>
		 (<%=oRec("jobnr")%>)
		<%end if %> 
		 </a>
		<%else%>
		<font class=medlillesilver><%=oRec("jobnavn")%> 
		<%if cint(t_skjuljobnr) <> 1 then %>
		 (<%=oRec("jobnr")%>)
		<%end if %> </font>
		<%end if%>
		&nbsp;(<%=oRec("antalakt")%>)
		
		<%if cint(t_sorter) <> 0 then %>
		<%if (cint(s) = 1 AND cint(t_lickid) = 0) OR cint(s) = 2 then %>
		<br /><%=left(oRec("kkundenavn"), 17)%> <font class=medlillesilver>(<%=oRec("kkundenr")%>)</font>
		<%end if %>
		<%end if %>
		
		</td>
		
		<td align=right width="16" style="padding:8px 2px 2px 2px; border-bottom:1px silver dashed;" class=lille>
		<%=strprojektgrpOK%>
		</td>
		
	</tr>
	
	
	
	<%
	xi = xi + 1
	x = x + 1
	lastKid = oRec("kid")
	lastJobslutdato = oRec("jobslutdato")
	oRec.movenext
	wend
	oRec.close
	
	
	next 's
	
	
	if x = 0 then%>
	<tr bgcolor="#ffffff">
		<td colspan=2 style="padding:2px;">(Ingen)</td>
	</tr>
	<%end if%>
	
	</table>
	<br>
	<input type="hidden" name="lastjob" id="lastjob" value="0">
	</form>
	

</div>



     



<%
end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





