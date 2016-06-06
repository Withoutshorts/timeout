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
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	'*********** Sætter sidens global variable ***********************************
	
	'* Hvis der vælges en anden bruger
	'* den den der er logget på.
	 
	 '*** Mid !***
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
	window.top.frames['a'].location.href = "timereg_akt_2006.asp?usemrn=<%=usemrn%>&showakt=0"
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
	<div id="sindhold" style="position:absolute; left:20; top:70; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu()%>
<%

  

 
 
 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	
				'*** Tilvalgte job i guiden dine aktive job ****
				varUseJob = "("
				varAktivejob = "("
				
				if len(request("FM_sog_job_navn_nr")) <> 0 then
				sogjobnavn_nr = request("FM_sog_job_navn_nr")
				show_sogjobnavn_nr = sogjobnavn_nr
				else
				sogjobnavn_nr = 0
				show_sogjobnavn_nr = ""
				end if
				
				if len(request("FM_kontakt")) <> 0 then
				sogKontakter = request("FM_kontakt")
				else
				sogKontakter = 0
				end if
				
				if len(request("FM_ignorer_projektgrupper")) <> 0 then
				ignorerProgrp = 1
				selIgn = "CHECKED"
				else
				ignorerProgrp = 0
				selIgn = ""
				end if
				
				'*****************************************
				'**** Søger på job ***********************
				'*** Kontakter "overruler" søg på job **** 
				'*****************************************
				
				if len(trim(show_sogjobnavn_nr)) <> 0 AND sogKontakter = 0 then
				
					strSQL3 = "SELECT id, jobnr, jobnavn FROM job WHERE (jobnr = '"& sogjobnavn_nr &"' OR jobnavn LIKE '"& sogjobnavn_nr &"%') AND jobstatus = 1"
					'Response.write strSQL3
					oRec3.open strSQL3, oConn, 3
					
					while not oRec3.EOF 
					varUseJob = varUseJob & " j.id = "& oRec3("id") & " OR "
					varAktivejob = varAktivejob & " jobid = "& oRec3("id") & " OR " 
					oRec3.movenext
					wend 
					
					oRec3.close 
				
				else
				
				
					'*** Er der søgt på kontakt? ***
					if sogKontakter <> 0 then
					
					strSQL = "SELECT id, jobknr FROM job WHERE jobknr = "& sogKontakter
				
					oRec.open strSQL, oConn, 3
					while not oRec.EOF
					varUseJob = varUseJob & " j.id = "& oRec("id") & " OR "
					varAktivejob = varAktivejob & " jobid = "& oRec("id") & " OR "
					oRec.movenext
					wend
					
					oRec.close
					
					end if
				
				end if
				
				'********************************
				
				
				
				
				'****************************************
				'*** Skal projektgrupper ignoreres ?? ***
				if ignorerProgrp = 0 then
					
					if varAktivejob = "(" then
					varAktivejob = ""
					else
					varAktivejob_len = len(varAktivejob)
					varAktivejob_left = varAktivejob_len - 3
					varAktivejob_use = left(varAktivejob, varAktivejob_left) 
					varAktivejob = varAktivejob_use & ") AND "
					end if
				
				else
				
					strPgrpSQLkri = ""
					varAktivejob = ""
					
				end if
				
				
				
				
				
				if len(trim(show_sogjobnavn_nr)) = 0 OR sogKontakter = 0 OR (len(trim(show_sogjobnavn_nr)) <> 0 AND ignorerProgrp = 0)  then
				
						'******************************************************
						'**** Henter job fra Guiden Dine aktive job ***********
						strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE "& varAktivejob &" medarb = "& usemrn &""
						oRec3.open strSQL3, oConn, 3
						
						while not oRec3.EOF 
						varUseJob = varUseJob & " j.id = "& oRec3("jobid") & " OR "
						oRec3.movenext
						wend 
						
						oRec3.close 
						
						
				end if
				
				
				
				if varUseJob = "(" then
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
	
	
	%>
	<div name="filter" id="filter" style="position:absolute; left:0px; top:30px; height:180px; display:; visibility:visible; background-color:#ffffe1; width:450px; padding:5px; border:1px #5582d2 solid; border-right:2px #5582d2 solid; border-bottom:2px #5582d2 solid;">
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<form action="timereg_2006.asp" name="sogekri" method="POST">
	
	<tr>
		<td><b>A) Jobnr</b> eller <b>jobnavn:</b></td>
		<td bgcolor="#ffffe1" valign=top> 
		<input type="text" name="FM_sog_job_navn_nr" id="FM_sog_job_navn_nr" value="<%=show_sogjobnavn_nr%>" style="font-family:arial; font-size:10px; width:220px;"> 
		</td>
	</tr>
	<tr>
	<%
	'**** Finder kunder til dd ****
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
	<td><b>B) Kontakt:</b></td>
	<td>
	<select name="FM_kontakt" style="font-family: arial; font-size: 10px; width:220px; height:12px;">
		<option value="0">Alle</option>
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
		
		</select>
			
	</td>
	
	</tr>
	<tr>
		<td colspan=2 style="padding-top:5px; padding-bottom:5px;">
		<input type="checkbox" name="FM_ignorer_projektgrupper" id="FM_ignorer_projektgrupper" value="1" <%=selIgn%>>Ignorer <b>projektgrupper</b>, og <b>"Guiden dine aktive job"</b> 
		&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="Søg"></td>
	</tr>
	</form>
	<form action="timereg_2006.asp?func=skiftmedarb" name="selmed" method="POST">
	<tr bgcolor="#ffffff">
	<td style="padding-top:5px; padding-bottom:5px; border-top:1px #003399 dashed; border-bottom:1px #003399 dashed;"><b>Indtast timer for:</b></td>
	<td style="padding-top:5px; padding-bottom:5px; border-top:1px #003399 dashed; border-bottom:1px #003399 dashed;"><select name="FM_use_me" style="font-family: arial; font-size: 10px; width:220px; height:12px;">
				<%
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere ORDER BY Mnavn"
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
	
	&nbsp;&nbsp;<input type="submit" value="Skift"></td>
	</tr>
	</form>
	</table><img src="../ill/blank.gif" width="1" height="5" alt="" border="0"><br>
	<a href="javascript:popUp('stopur_2006.asp?FM_use_me=<%=usemrn%>','600','500','250','120');" target="_self"; class=vmenu><img src="../ill/stopur.gif" width="21" height="18" alt="" border="0"> Stopur. (hjælp til timereigstrering)</a>
	<br>
	<img src="../ill/ac0052-24.gif" width="24" height="24" alt="" border="0">&nbsp;<a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>','600','500','250','120');" target="_self"; class=vmenu>Se "Guiden dine aktive job" (Skjulte job)</a>
	</div>
	
	
	
	<div name="kontakterogjob" id="kontakterogjob" style="position:absolute; left:460px; top:30px; display:; width:500px; visibility:visible; height:180px; background-color:#ffffff; overflow:auto; padding:5px; border:1px #5582d2 solid; border-right:2px #5582d2 solid; border-bottom:2px #5582d2 solid;">
	
	<table cellspacing=0 cellpadding=0 border=0 width=470>
	<form>
	<tr bgcolor="#5582d2">
		<td height=20 class=alt style="padding-left:5px;"><b>Adg?</b></td>
		<td width=230 class=alt style="padding-left:3px;"><b>Job</b> (Aktiviter)</td>
		<td class=alt width=30><b>Skjul*</b>&nbsp;&nbsp;</td>
		<td class=alt>&nbsp;<b>Mat.</b>&nbsp;</td>
		<td class=alt><b>Timer</b></td>
	</tr>
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
	
	
	'************** Main SQL Call ******************************************
	strSQL = "SELECT j.id AS id, jobnr, jobnavn, jobknr, kkundenavn, kkundenr, count(a.id) AS antalakt,"_
	&" j.budgettimer, j.ikkebudgettimer, "_
	&" kid, jobans1, jobans2 FROM job j, kunder k"_
	&" LEFT JOIN aktiviteter a ON (a.job = j.id)"
	
	if ignorerProgrp = 0 then
	strSQL = strSQL &" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 AND k.Kid = j.jobknr "& strPgrpSQLkri & " GROUP BY j.id, a.job ORDER BY k.kkundenavn, j.jobnavn, j.id, j.jobnavn" 
	else
	strSQL = strSQL &" WHERE "& varUseJob &" j.jobstatus = 1 AND j.fakturerbart = 1 AND k.Kid = j.jobknr GROUP BY j.id, a.job ORDER BY k.kkundenavn, j.jobnavn, j.id, j.jobnavn" 
	end if
	
	'Response.write strSQL
	'Response.flush
	'Response.write "<hr>"
	
	
	
	oRec.Open strSQL, oConn, 3
	
	lastKid = 0
	x = 0
	while not oRec.EOF 
	
	editok = 0
	bgThis = "ffffff"
	
	
	'** rettigheder ** 
		if level = 1 then
		editok = 1
		else
				if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
				editok = 1
				end if
		end if
	
	if lastKid <> oRec("kid") then%>
	<tr bgcolor="#8caae6">
		<td colspan=6 height=20 class=alt style="padding-left:5px; border-top:1px #5582d2 solid; border-bottom:1px silver solid;"><b><%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)</b></td>
	</tr>
	<%end if%>
	
	<tr bgcolor="<%=bgThis%>" id="tr_<%=oRec("id")%>">
		<td align=center style="border-bottom:1px silver solid;">
		<%
		if (instr(strJobmedadgang, "#"& oRec("id") &"#") <> 0) OR ignorerProgrp = 0 then
		projektgrpOK = 1
		
			%>
			<font color="green"><i>V</i></font>
			
		<%
		
		else
		projektgrpOK = 0
			if editok = 1 then%>
			<a href="javascript:popUp('tilknytprojektgrupper.asp?id=<%=oRec("id")%>&medid=<%=usemrn%>','600','500','150','120');" target="_self"; class=vmenu><font color="red">X</font></a>
			<%else
			%>
			<font color="red">X</font>
			<%end if
		end if
		
		%>
		</td>
		
		<td style="padding-left:3px; border-bottom:1px silver solid;">
		<%if projektgrpOK = 1 then%>
		<a href="timereg_akt_2006.asp?jobid=<%=oRec("id")%>&usemrn=<%=usemrn%>&showakt=1" target="a" onClick=markerjob(<%=oRec("id")%>); class=vmenu><img src="ill/plus.gif" width="9" height="9" alt="" border="0">&nbsp;&nbsp;
		<%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</a>
		<%end if%>
		
		<%if editok = 1 then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=treg" target="_top">
		<img src="../ill/ac0059-16.gif" width="16" height="16" alt="" border="0">
		</a>
		<%else%>
		<font class='storsilver'><b><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</b></font>
		<%end if%>
		&nbsp;&nbsp;&nbsp;
		<b>(<%=oRec("antalakt")%>)</b></td>
		
		
		<td align=center style="border-bottom:1px silver solid;"><input type="checkbox" name="FM_flyttilguide" value="<%=oRec("id")%>"></td>
		<td style="border-bottom:1px silver solid; padding-top:2px;">
		<%'if lto = "dencker" OR lto = "demo" OR lto = "intranet" then%>
		<%'end if%>
		
		<%'if projektgrpOK = 1 then%>
		&nbsp;&nbsp;<a href="javascript:popUp('materialer_indtast.asp?id=<%=oRec("id")%>','600','620','150','20');" target="_self" class=alt><img src="../ill/ac0038-16.gif" width="16" height="16" border="0" alt="Materiale forbrug / udlæg"></a>
		<%'end if%>&nbsp;
		</td>
		<td style="border-bottom:1px silver solid; padding-top:2px;">
		<%if projektgrpOK = 1 then%>
		<a href="timereg_akt_2006.asp?jobid=<%=oRec("id")%>&usemrn=<%=usemrn%>&showakt=1" target="a" onClick=markerjob(<%=oRec("id")%>);><img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		<%end if%>&nbsp;
		</td>
	</tr>
	
	
	
	<%
	x = x + 1
	lastKid = oRec("kid")
	oRec.movenext
	wend
	oRec.close
	
	%>
	<tr>
		<td colspan=4 align=right style="padding:5px;"><input type="submit" value="*Flyt til Guiden aktive job. (Skjul)"><br>&nbsp;</td>
	</tr>
	</table>
	<br>
	<input type="hidden" name="lastjob" id="lastjob" value="0">
	</form>
	

</div>


<%
'*********************************************************
'***** Stempelur *****************************************
'*********************************************************
if session("stempelur") <> 0 then

	strSQL2 = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud "_
	&" FROM login_historik l"_
	&" WHERE l.logud IS NULL AND l.mid = " & useMrn &""_
	&" ORDER BY l.login DESC"
	
	loginhist = 0
	strLoginHist = ""
	oRec2.open strSQL2, oConn, 3 
	while not oRec2.EOF 
	loginhist = 1
	strLoginHist = strLoginHist & oRec2("login") & "<br>"
	
	oRec2.movenext
	wend
	oRec2.close  
	
if loginhist <> 0 AND session("stempelur_vist") <> 1 then
vzbLogHist = "visible"
dspLogHist = ""
session("stempelur_vist") = 1
else
vzbLogHist = "hidden"
dspLogHist = "none"
end if

if loginhist = 0 then
strLoginHist = "(Ingen!)<br><br>"
end if

%>

<div name=stempelur id=stempelur style="position:absolute; left:460; top:30; visibility:<%=vzbLogHist%>; display:<%=dspLogHist%>; width:500; height:200px; border:1px #000000 solid; border-right:2px #000000 solid;  border-bottom:2px #000000 solid; overflow:auto; background-color:#ffffff; filter:alpha(opacity=90); z-index:10000; padding:10px;">
<h3>Login Historik (Stempelur)</h3>
<a href="javascript:popUp('stempelur.asp?func=stat&medarbSel=<%=useMrn%>&showonlyone=1&hidemenu=1','700','500','200','120');" target="_self"; class=vmenu>[Angiv logud-tider]</a>
<img src="../ill/blank.gif" width="20" height="14" alt="" border="0"><a onClick="lukstempelur()" href="#" class="red">[X]</a><br><br>

<b>Du har følgende uafsluttede logins:</b><br>
<%=strLoginHist%><br>

<a href="javascript:popUp('stempelur.asp?func=stat&medarbSel=<%=useMrn%>&showonlyone=1&hidemenu=1','700','500','200','120');" target="_self"; class=vmenu>klik her </a> for at angive logud tider på ovenstående logins.<br><br>
<b>NB)</b><br>
Husk altid at afslutte din dag ved at logge ud<br>
på den røde logud knap <img src="../ill/logud.gif" width="46" height="14" alt="" border="0"> øverst til højre i TimeOut.<br><br>&nbsp;

<%
'****************** Stempelur slut ******
%>
</div>

&nbsp;|&nbsp;&nbsp;<a href="#" onClick=visstempelur(); class=rmenu>Vis Login Historik</a>
<%end if%>


<%
end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





