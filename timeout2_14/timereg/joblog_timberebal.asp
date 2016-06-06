<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->



<%

func = request("func")
thisfile = "joblog_timberebal"

	if len(request("FM_medarb")) <> 0 then
	FM_medarb = request("FM_medarb")
	else
	FM_medarb = request("FM_medarb_alle")
	end if
	
	'**** Valgte medarbejdere ***
	sqlMedKri = "("
	antalmedarb = 0
	medarbs = split(FM_medarb, ",")
	for i = 0 to UBOUND(medarbs)
	sqlMedKri = sqlMedKri & " m.mid = "& medarbs(i) &" OR "
	antalmedarb = antalmedarb + 1
	next
	
	len_sqlMedKri = len(sqlMedKri) 
	left_sqlMedKri = left(sqlMedKri, len_sqlMedKri-3)
	sqlMedKri = left_sqlMedKri & ")"
	
	'** Alt medarbSQL kriterie
	sqlMedKriThis = replace(sqlMedKri, "m.mid", "tmnr")
												
	
	if len(request("FM_job")) <> 0 then
	FM_job = request("FM_job")
	else
	FM_job = 0
	end if
	
	'**** Valgte job ****
	sqlJobKri = "("
	
	jobnrs = split(FM_job, ",")
	antaljob = 0
	for i = 0 to UBOUND(jobnrs)
	sqlJobKri = sqlJobKri & " jobnr = "& jobnrs(i) &" OR "
	antaljob = antaljob + 1
	next
	
	len_sqlJobKri = len(sqlJobKri) 
	left_sqlJobKri = left(sqlJobKri, len_sqlJobKri-3)
	sqlJobKri = left_sqlJobKri & ")"
	
	
	'*** Hvad skal vises ***
	'0 timer
	'1 fakbare timer og oms
	'2 ressource timer
	
	if len(request("vis_fakbare_res")) <> 0 then
	visfakbare_res = request("vis_fakbare_res")
	response.cookies("cc_vis_fakbare_res") = visfakbare_res
	response.cookies("cc_vis_fakbare_res").expires = date + 10
	else
		if len(request.cookies("cc_vis_fakbare_res")) <> 0 then
		visfakbare_res = request.cookies("cc_vis_fakbare_res")
		else
		visfakbare_res = 0
		end if
	end if
	
	if (antaljob <= 25 OR (antalmedarb <= 5 AND antaljob <= 150)) then
	visfakbare_res = visfakbare_res
	else
		if visfakbare_res = 2 then
	 	visfakbare_res = 0
	 	else
		visfakbare_res = visfakbare_res
	 	end if
	end if
	
	
	'if len(request("vis_medarb")) <> 0 then
	'visMedarb = request("vis_medarb")
	'response.cookies("cc_vis_medarb") = visMedarb
	'response.cookies("cc_vis_medarb").expires = date + 10
	'else
	'	if request.cookies("cc_vis_medarb") = "1" then
	'	visMedarb = 1
	'	else
	'	visMedarb = 0
	'	end if
	'end if
	
	if len(request("vis_job")) <> 0 then
	visJob = request("vis_job")
	response.cookies("cc_vis_job") = visJob
	response.cookies("cc_vis_job").expires = date + 10
	else
		if request.cookies("cc_vis_job") = "1" then
		visJob = 1
		else
		visJob = 0
		end if
	end if
	
	
	
	if len(request("FM_vis_aktiviteter_forjob")) <> 0 then
	forjob = request("FM_vis_aktiviteter_forjob")
	else
	forjob = 0
	end if
	
	'if request("vis_jobtotalpavalgtemedarb") = 1 then
	visTotaluansetMedarb = 1
	'else
	'visTotaluansetMedarb = 0
	'end if
	
	
	'*** txt orientering ***
	'if request("FM_orientering") = "v" OR request.cookies("cc_orientering") = "v" then
		tdstyleTimOms = "valign=top style='padding:2px; border-left:1px #999999 solid; border-top:1px #999999 solid; writing-mode:tb-rl;'"
		tdstyleTimOms1 = "valign=top style='padding:2px; border-left:2px #999999 solid; border-top:1px #999999 solid; writing-mode:tb-rl;'"
		oriChk1 = "CHECKED"
		oriChk = ""
		tdh = "250"
	
	'else
		
		tdstyleTimOms2 = "valign=bottom style='padding:2px; border-left:1px #999999 dashed; border-top:1px #999999 solid;'"
		tdstyleTimOms3 = "valign=bottom style='padding:2px; border-left:2px #999999 solid; border-top:1px #999999 solid;'"
	
	'	oriChk = "CHECKED"
	'	oriChk1 = ""
	
	'tdh = "20"
	'end if
	
	if len(request("FM_orientering")) <> 0 then
	response.cookies("cc_orientering") = request("FM_orientering")
	response.cookies("cc_orientering").expires = date + 10
	end if
	
	'** Vis nulfilter ***
	if len(request("FM_visnulfilter")) <> 0 then
	visnulfilter = 1
	nulFilter = "CHECKED"
	else
	visnulfilter = 0
	nulFilter = ""
	end if
	
	
	public restimerThis, strJobLinie
	function restimer(medi, joid, stdato, sldato)
		
		
		strSQL3 = "SELECT sum(r.timer) AS restimer FROM "_
		&" ressourcer r WHERE r.mid = "& medi &" AND r.jobid = "& joid &" "_
		&" AND r.dato BETWEEN '"& stdato &"' AND '"& sldato &"'"_
		&" GROUP BY r.jobid"
		
		'Response.write strSQL3
		'Response.flush
		
		restimerThis = 0
		oRec3.open strSQL3, oConn, 3 
		if not oRec3.EOF then 
		restimerThis = oRec3("restimer")
		end if
		oRec3.close
		
		if restimerThis > 0 then
		strJobLinie = strJobLinie & formatnumber(restimerThis, 2)
		else
		strJobLinie = strJobLinie & "&nbsp;"
		end if
		
	end function
	
	public intKorselKM '***( Bruges ikke endnu)
	function korsel(medi, jonr, stdato, sldato)
			
			
			strSQL3 = "SELECT sum(t.timer) AS sumtimerFakbare, tmnr, tjobnr, sum(t.timer*t.timepris) AS vaegtettimepris "_
			&" FROM medarbejdere m "_
			&" LEFT JOIN timer t ON tjobnr = "& jonr &" AND tmnr = m.mid AND tfaktim = 5 AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut&"'"_
			&" WHERE m.mid = "& medi &" GROUP BY tmnr, tjobnr ORDER BY tmnr"
			
			
			'Response.write strSQL3
			'Response.flush
			
			intKorselKM = 0
			oRec3.open strSQL3, oConn, 3 
			if not oRec3.EOF then 
			intKorselKM = oRec3("sumtimerFakbare")
			end if
			oRec3.close
			
			if intKorselKM > 0 then
			Response.write formatnumber(intKorselKM, 2)
			else
			Response.write "&nbsp;"
			end if
	
	end function
	
	
	
	
	
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
	
	<SCRIPT Language=JavaScript>
	function BreakItUp()
	{
	  //Set the limit for field size.
	  var FormLimit = 102399
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	</SCRIPT>

	
	
	
	
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
	
	
	
	
	<div id="sindhold" style="position:absolute; left:20px; top:132px; visibility:visible;">
	<table>
	<tr></tr>
	<td Style="border:1px #8caae6 solid; border-bottom:2px #8caae6 solid; border-right:2px #8caae6 solid; padding:10px;" bgcolor="#ffffff">
	<table>
	<form action="joblog_timberebal.asp" method="post">
	<input type="hidden" name="FM_medarb" id="FM_medarb" value="<%=FM_medarb%>">
	<input type="hidden" name="FM_job" id="FM_job" value="<%=FM_job%>">
		<tr>
			<td valign=top width=350>
			<h3>Filter og søgekriterier.</h3>
			<b>Periode:</b><br><!--#include file="inc/weekselector_s.asp"--></td>
			<td rowspan=2 valign=top style="padding-top:5px;" width=350>
			<!--<b>Vis / udspecificer:</b>-->
			<table cellpadding=2 cellspacing=2 border=0>
			<tr>
				<td colspan=2 valign=top style="padding-top:5px;" bgcolor="#d6dff5"><b>Vis Totaler uanset valgte medarbejder(e):</b><br>
				<font class=megetlillesort>Søjler med timefordeling pr. medarbejder vises ikke. 
				<br>Jobtotaler vises som total uanset valgte medarbejdere
				</td>
				
				<%if cint(visJob) = 1 then
				jbChk1 = "CHECKED"
				jbChk = ""
				else
				jbChk1 = ""
				jbChk = "CHECKED"
				end if%>
				
				<td valign=top bgcolor="#d6dff5"><input type="radio" name="vis_job" id="vis_job" value="1" <%=jbChk1%>></td>
			</tr>
			<tr>
				<td colspan=2 valign=top style="padding-top:5px;" bgcolor="#d6dff5"><b>Udspecificer på de valgte medarbejder(e):</b><br>
				<font class=megetlillesort>Jobtotaler beregnes udfra de valgte medarbejdere.<br>
				Timer udspecificeres pr. medarbejder
				</td>
				<td valign=top bgcolor="#d6dff5"><input type="radio" name="vis_job" id="vis_job" value="0" <%=jbChk%>></td>
			</tr>
			<tr>
				<td colspan=3 valign=top bgcolor="#ffffff"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			</tr>
			
			<%select case cint(visfakbare_res) 
			case 1
				
				fakChk = ""
				fakChk1 = "CHECKED"
				fakChk2 = ""
				
			case 2
				
				fakChk = ""
				fakChk1 = ""
				fakChk2 = "CHECKED"
				
			case else
				
				fakChk = "CHECKED"
				fakChk1 = ""
				fakChk2 = ""
			
			end select%>
			
			<tr>
				<td colspan=2 valign=top style="padding-top:5px;" bgcolor="#eff3ff"><b>Vis kun timer.<br>
				<font class=megetlillesort>Viser alle fakturerbar og ikke fakturerbare timer.</b></td>
				<td valign=top style="padding-top:5px;" bgcolor="#eff3ff"><input type="radio" name="vis_fakbare_res" id="vis_fakbare_res" value="0" <%=fakChk%>></td>
			</tr>
			
			<tr>
				<td colspan=2 valign=top style="padding-top:5px;" bgcolor="#ffffff"><b>Vis fakturerbare timer og omsætning.</b><br>
				<font class=megetlillesort>Ved udsp. af aktiviteter vises også ikke fakturerbare akt.</td>
				<td valign=top style="padding-top:5px;" bgcolor="#ffffff"><input type="radio" name="vis_fakbare_res" id="vis_fakbare_res" value="1" <%=fakChk1%>></td>
			</tr>
			<%if antaljob <= 25 OR (antalmedarb <= 5 AND antaljob <= 150) then%>
			<tr>
				<td colspan=2 valign=top bgcolor="#ffff99" style="padding-top:5px;"><b>Vis registrerede timer, ressourcetimer  og balance:</b><br>
				<font class=megetlillesort>Vises ikke ved udspe. af aktiviteter</font></td>
				<td valign=top bgcolor="#ffff99"><input type="radio" name="vis_fakbare_res" id="vis_fakbare_res" value="2" <%=fakChk2%>></td>
			</tr>
			<%else%>
			<tr>
				<td colspan=3 valign=top bgcolor="#ffff99" style="padding-top:5px;">
				<font class=megetlillesort>Ressourcetimer, Udspecificering kan kun ske hvis:<br>
				- der er valgt mindre end 25 job.<br>
				- der er valgt mindre end 150 job og maks 5 medarbejder.</td>
			</tr>
			<%end if%>
			
			</table></td>
		</tr>
		<tr>
			<td style="padding-top:5px;">
				<%
				'if visTotaluansetMedarb = 1 then
				'chkMedJob1 = "CHECKED"
				'chkMedJob = ""
				'else
				'chkMedJob1 = ""
				'chkMedJob = "CHECKED"
				'end if%>
				
				<!--<input type="radio" name="vis_jobtotalpavalgtemedarb" id="vis_jobtotalpavalgtemedarb" value="1" <=chkMedJob1%>> Vis totaltimer på job som sum af de <u>valgte</u> medarbejdere<br>
				<input type="radio" name="vis_jobtotalpavalgtemedarb" id="vis_jobtotalpavalgtemedarb" value="0" <=chkMedJob%>> Vis totaltimer på job <u>uanset</u> hvilke medarbejdere der er valgt.<br><br>-->
				
				<b>Vis "NUL" job (job uden timeregistreringer) </b><br>
				Ja: <input type="checkbox" name="FM_visnulfilter" id="FM_visnulfilter" value="1" <%=nulFilter%>>  
				<br><br>
			
				Udspecificér <b>Aktiviteter</b> på et enkelt job: (vælg job)<br>
				<select name="FM_vis_aktiviteter_forjob" id="FM_vis_aktiviteter_forjob" style="width:300px;">
				<option value=0>Vis alle job</option>
				<option value=0>Vis aktiviter (Vælg job)</option>
				<option value=0>-------------------------------------------------------------</option>
				<%
				strSQLjob = "SELECT id, jobnr, jobnavn"_
				&" FROM job "_
				&" WHERE "& sqlJobKri &" ORDER BY jobnavn"
				
				oRec.open strSQLjob, oConn, 3 
				while not oRec.EOF 
				if cint(forjob) = oRec("jobnr") then
				thisSEL = "SELECTED"
				else
				thisSEL = ""
				end if%>
				<option value="<%=oRec("jobnr")%>" <%=thisSEL%>><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</option>"
				<%
				oRec.movenext
				wend
				oRec.close 
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
			<!--Text-orientering på medarbejdere timer og beløb:<br><input type="radio" name="FM_orientering" value="h" <=oriChk%>> <b>Horisontal</b>  &nbsp;&nbsp;
			<input type="radio" name="FM_orientering" value="v" <=oriChk1%>>&nbsp;<b>Vertical</b> (til print)-->
			&nbsp;</td>
			<td align=right><input type="image" src="../ill/statpil.gif"></td>
		</tr>
	</form>
	</table>
	</td></tr>
	</table>
	
	
	</div>
	
	

	
	<div id="antal" style="position:absolute; left:762px; top:137px; visibility:visible; background-color=#ffffff; border:1px #8caae6 solid; border-bottom:2px #8caae6 solid; border-right:2px #8caae6 solid; padding:10px;">
	<%
	'jur = 0
	
	if (antalmedarb + antaljob) < 25 then
	faktor = (antalmedarb + antaljob) / 1.5
	end if
	
	if (antalmedarb + antaljob) >= 25 AND (antalmedarb + antaljob) < 50  then
	faktor = (antalmedarb + antaljob) / 2
	end if
	
	if (antalmedarb + antaljob) >= 50 AND (antalmedarb + antaljob) < 75  then
	faktor = (antalmedarb + antaljob) / 5
	end if
	
	if (antalmedarb + antaljob) >= 75 AND (antalmedarb + antaljob) < 100  then
	faktor = (antalmedarb + antaljob) / 10
	end if
	
	if (antalmedarb + antaljob) >= 100 then
	faktor = (antalmedarb + antaljob) / 18
	end if
	
	if cint(visJob) = 1 then
	faktor = faktor * 6
	end if
	
	if visfakbare_res = 2 then
	faktor = faktor / 6
	end if
	
	Response.write "<h3>Loadtid</h3>"
	Response.write "Der er valgt: <b>" & antalmedarb & "</b> medarbejder(e).<br>"
	Response.write "Der er valgt: <b>" & antaljob & "</b> job.<br><br>"
	'Response.write faktor
	Response.write "<b><u>Forventet loadtid:</u></b><br><font color=red> < "& cint((antalmedarb + antaljob) / faktor) &" Sekunde(r).</font> (2 mbit ADSL)"
	Response.flush
	%><br><br>
	<a href="stat.asp?menu=stat&show=joblog_timberebal" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage til søgekriterier</a></div>
	
	
	<%
	if (antaljob) > 250 then
	%>
	<div id="fejl" style="position:absolute; left:200px; top:360px; visibility:visible; border-color:2px red solid; background-color=#ffff99; border:2px red solid; padding:10px;">
	<h3><img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">Fejl!</h3>
	Antallet af valgte job er over 250.<br>
	Vælg et mindre antal.<br><br>
	<a href="Javascript:history.back()"><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a>
	</div>
	<%
	else
	%>	
	<div id="oCode" name="oCode" style="position:absolute; zoom: 100%; left:20px; top:400px; visibility:visible;" bgcolor="#ffffff">
	<%	
			m = 20
			x = 10500 '6500 '2500
			lastjid = 0
			strMids = 0
			
			
			dim jobmedtimer ', jobtimerIaltX
			redim jobmedtimer(x, m) ', jobtimerIaltX(x)
			
			v = 0
			j = 0
			dim medarb
			redim medarb(v)
			
			dim medabTottimer, restimerTot, omsTot, fakbareTimerTot, normTimerTot
			Redim medabTottimer(v), restimerTot(v), omsTot(v), fakbareTimerTot(v), normTimerTot(v)
			
			dim  jobTottimer, jobRestimerTot, jobOmsTot, jobfakbareTimerTot
			Redim  jobTottimer(0), jobRestimerTot(0), jobOmsTot(0), jobfakbareTimerTot(0)
			
			
			'*** Udspecificer på aktiviteter? **'
			if forjob <> 0 then
				
				if cint(visJob) = 1 then
				grpBySQL = "a.id"
				orderBySQL = "a.id"
				else
				grpBySQL = "a.id, m.mid"
				orderBySQL = "a.id, m.mid"
				end if
			
			
			sqlJobKri = " j.jobnr = "& forjob
			aktSQLFields = ", taktivitetid, taktivitetnavn, a.id AS aid, a.job, a.navn AS anavn, a.fakturerbar AS afakbar, a.budgettimer AS abudgettimer, a.aktbudget AS abudget "
			aktSQLTables = ", aktiviteter a "
			aktSQLWhere = " AND (a.job = j.id AND a.fakturerbar <> 2) "
			aktjobTimerWhere = " taktivitetid = a.id "
			else
				
				
				if cint(visJob) = 1 then
				grpBySQL = "jobnr"
				orderBySQL = "jobnr"
				else
				grpBySQL = "jobnr, m.mid"
				orderBySQL = "jobnr, m.mid"
				end if
			
			
			sqlJobKri = sqlJobKri
			aktSQLFields = ""
			aktSQLTables = ""
			aktSQLWhere = ""
			aktjobTimerWhere = " tjobnr = j.jobnr "
			end if
			
			if visfakbare_res = 1 then
				if forjob = 0 then
				whrClaus = " t.tfaktim = 1"
				else
				whrClaus = " (t.tfaktim = 1 OR tfaktim = 2) " '** aktiviteter
				end if
			vgtTimePris = " sum(t.timer*t.timepris) AS vaegtettimepris, "
			else
			whrClaus = " (t.tfaktim = 1 OR tfaktim = 2) "
			vgtTimePris = ""
			end if
			
			sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
			sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
			
			if cint(visJob) = 1 then
			sqlMedKri = "(m.mid <> 0)"
			else
			sqlMedKri = sqlMedKri
			end if
			
			
			
			'*** Main SQL ******
			strSQL = "SELECT j.id AS jid, t.tjobnavn, t.tjobnr, j.jobnr AS jnr, j.jobnavn, "_
			&" sum(t.timer) AS sumtimer, "& vgtTimePris &" t.tmnr, t.tmnavn, t.timepris, m.mnavn, m.mnr, m.mid, "_
			&" j.jobTpris, j.fakturerbart, j.budgettimer, j.ikkebudgettimer, j.fastpris, m.medarbejdertype AS mtype "& aktSQLFields &","_
			&" k.kkundenavn, k.kkundenr"_
			&" FROM medarbejdere m, job j "& aktSQLTables &""_
			&" LEFT JOIN kunder k ON (k.kid = j.jobknr)"_
			&" LEFT JOIN timer t ON ( "& aktjobTimerWhere &" AND tmnr = m.mid AND "& whrClaus &" AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut&"')"_
			&" WHERE "& sqlMedKri &" AND "& sqlJobKri &" "& aktSQLWhere &" GROUP BY "& grpbySQL &" ORDER BY " & orderBySQL
			
			
			'Response.write "<br><br><br>"
			'response.write strSQL & "<br><br>"
			'response.flush
			
			x = 0
			k = 0
			timeprisThis = 0
			strJobnbs = "#0#,"
			strMidsK = "#0#,"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
						
						
						
						'Response.write x &", " & m &"<br>"
						
						'*** Job og medarb data ****
						jobmedtimer(x, 0) = oRec("jid")
						jobmedtimer(x, 1) = oRec("jobnavn")
						jobmedtimer(x, 3) = formatnumber(oRec("sumtimer"), 2)
						jobmedtimer(x, 2) = oRec("mnavn") &" ("&oRec("mnr")&")"
						jobmedtimer(x, 4) = oRec("mid")
						
						'jobmedtimer(x, 5) = oRec("restimer")
						jobmedtimer(x, 6) = oRec("jnr")
						
						'*** Uspecificering af aktiviteter ***
						if cint(forjob) <> 0 then
							
							jobmedtimer(x, 12) = oRec("aid")
							jobmedtimer(x, 13) = oRec("anavn")
							thisAktFakbar = oRec("afakbar")
							if thisAktFakbar = 1 then
							jobmedtimer(x, 14) = "<font color=green>Fakturerbar"
							else
							jobmedtimer(x, 14) = "<font color=green>Ikke fakturerbar"
							end if
						
						end if
						
						jobmedtimer(x,15) = oRec("mtype")
						
						
						'*** Antal Medarbejdere / Navne ***
						if instr(strMidsK, "#"& oRec("mnr") &"#") = 0 then
						strMidsK = strMidsK & ",#"& oRec("mnr") &"#"
							
							Redim preserve medarb(v)
							Redim preserve medarbnavnognr(v)
							medarb(v) = jobmedtimer(x,4)
							medarbnavnognr(v) = jobmedtimer(x, 2)
							v = v + 1
							
						else
						strMidsK = strMidsK
						end if
						
						
						'*** Jobtype ***
						if oRec("fastpris") = 1 then
						jobmedtimer(x, 8) = "<font color=green>Fastpris"
						else
						jobmedtimer(x, 8) = "<font color=green>Lbn.timer" 
						end if
						
						
						'** Timepris v. fastpris job 
						timeprisThis = 0
						
						if oRec("budgettimer") > 0 then
						bdgTim = oRec("budgettimer")
						else
						bdgTim = 1
						end if
						
						if oRec("jobtpris") > 0 then
						jobbelob = oRec("jobtpris")
						else
						jobbelob = 0
						end if
						
						if cint(forjob) <> 0 then
						'** budget **
						jobmedtimer(x, 18) = oRec("abudget")
						'** timer **
						jobmedtimer(x, 19) = oRec("abudgettimer")
						else
						'** budget **
						jobmedtimer(x, 18) = jobbelob
						'** timer **
							if cint(visfakbare_res) = 1 then
							jobmedtimer(x, 19) = bdgTim
							else
							jobmedtimer(x, 19) = bdgTim + oRec("ikkebudgettimer")
							end if
						end if
						
						timeprisThis = (jobbelob / bdgTim)
						
						if visfakbare_res = 1 then
							
							'*** Timepriser ***
							jobmedtimer(x,7) = 0
							if oRec("fastpris") = 1 then
							jobmedtimer(x,7) = (timeprisThis * oRec("sumtimer"))
							else
							jobmedtimer(x,7) = oRec("vaegtettimepris")
							end if
							
						end if
						
						jobmedtimer(x, 10) = timeprisThis
						jobmedtimer(x, 11) = oRec("fastpris")
						
							
							if forjob <> 0 then
							jora_id = oRec("aid")
							else
							jora_id = oRec("jnr")
							end if
							
							'*** Sumtimer på job ialt ***
							if instr(strJobnbs, "#"& jora_id &"#") = 0 then
							strJobnbs = strJobnbs & ",#"& jora_id &"#"
								
								if len(oRec("sumtimer")) <> 0 then
								jobmedtimer(x,16) = oRec("sumtimer")
									if visfakbare_res = 1 then
									jobmedtimer(x,17) = jobmedtimer(x,7) 'oRec("vaegtettimepris")
									end if
								else
								jobmedtimer(x,16) = 0
									if visfakbare_res = 1 then
									jobmedtimer(x,17) = 0
									end if
								end if
								lastj = x
								
								
							else
							jobmedtimer(lastj,16) = (jobmedtimer(lastj,16) + oRec("sumtimer"))
								if visfakbare_res = 1 then
									if oRec("fastpris") = 1 then
									jobmedtimer(lastj,17) = jobmedtimer(lastj,17)
									else
									jobmedtimer(lastj,17) = (jobmedtimer(lastj,17) + jobmedtimer(x,7))
									end if
								end if
							end if
						
						
							jobmedtimer(x, 20) = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"
						
						
				
			x = x + 1
			response.flush		
			oRec.movenext
			wend
			oRec.close 
			
			
			
			if cint(forjob) = 0 then
			jobaktOskrift = "Job"
			bgtd = "#ffffff"
			else
			jobaktOskrift = "Aktivitet"
			bgtd = "#ffffff"
			end if
			
			'*****************************************************************
			'*** Beregner antal mandage, tirsdag onsdag mv. ***
			'*** Bruges til at finde normerede timer i periode **
			wdayFirst = weekday(strDag&"/"&strMrd&"/"&strAar, 2)
			'Response.write wdayFirst & "<br>"
			
			wdayEnd = weekday(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2)
			'Response.write wdayEnd & "<br>"
			
			antaluger = datediff("ww", strDag&"/"&strMrd&"/"&strAar, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2, 3)
			'Response.write antaluger
			
			if 1 > wdayEnd AND 1 < wdayFirst then
			antalMandage = antaluger - 1
			else
			antalMandage = antaluger 
			end if
			
			if 2 > wdayEnd AND 2 < wdayFirst then
			antalTirsdage = antaluger - 1
			else
			antalTirsdage = antaluger 
			end if
			
			if 3 > wdayEnd AND 3 < wdayFirst then
			antalOnsdage = antaluger - 1
			else
			antalOnsdage = antaluger 
			end if
			
			if 4 > wdayEnd AND 4 < wdayFirst then
			antalTorsdage = antaluger - 1
			else
			antalTorsdage = antaluger 
			end if
			
			if 5 > wdayEnd AND 5 < wdayFirst then
			antalFredage = antaluger - 1
			else
			antalFredage = antaluger 
			end if
			
			if 6 > wdayEnd AND 6 < wdayFirst then
			antalLordage = antaluger - 1
			else
			antalLordage = antaluger 
			end if 
			
			if 7 > wdayEnd AND 7 < wdayFirst then
			antalSondage = antaluger - 1
			else
			antalSondage = antaluger 
			end if
			
			'************'			
			
			strJobLinie_top = ""
			strJobLinie = ""
			strJobLinie_total = ""
			
			
			strJobLinie_top = "<br><br><table cellspacing=0 cellpadding=0 border=0>"
			strJobLinie_top = strJobLinie_top & "<tr><td Style='border:1px #8caae6 solid; border-bottom:2px #8caae6 solid; border-right:2px #8caae6 solid; padding:10px;' bgcolor='#ffffff'>"
			strJobLinie_top = strJobLinie_top & "<h3>Timeforbrug, Omsætning og Ressourcetimer.</h3>"
			
			if cint(visfakbare_res) = 1 then
			strJobLinie_top = strJobLinie_top & "Realiserede fakturerbare timer og omsætning."
			else
			strJobLinie_top = strJobLinie_top & "Realiserede timer ialt. (Fakturerbare og ikke fakturerbare)"
			end if
			
			strJobLinie_top = strJobLinie_top & "<table cellspacing=0 cellpadding=0 border=0 bgcolor='#ffffff'>"
			strJobLinie_top = strJobLinie_top & "<tr>"
			strJobLinie_top = strJobLinie_top & "<td height="& tdh &" valign=bottom style='padding:2px; border-top:1px #999999 solid;' bgcolor='#d6dff5'><b>"&jobaktOskrift&" / Medarbejdere:</b><img src='../ill/blank.gif' width='250' height='1' alt='' border='0'></td>"
				
				
				
				'*** job totaler oskrifter  **
				
						if cint(forjob) <> 0 then
						strFakbtimTxt = "Forkalk. timer.&nbsp;"
						strFakbTxt = "Værdi. (Budget)&nbsp;"
						else
							if cint(visfakbare_res) = 1 then
							strFakbtimTxt = "Forkalkuleret fakturerbare timer.&nbsp;"
							else
							strFakbtimTxt = "Forkalk. timer ialt.&nbsp;"
							end if
							
							strFakbTxt = "Værdi. (Budget)&nbsp;"
						end if
						
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>"& strFakbtimTxt &"</td>"
				
				if cint(visfakbare_res) = 1 then
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>"& strFakbTxt &"</td>"
				end if
				
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>"&jobaktOskrift&"&nbsp;"
				
				if cint(visfakbare_res) = 1 AND forjob = 0 then
				strJobLinie_top = strJobLinie_top & "realiseret fak.bare timer"
				else
				strJobLinie_top = strJobLinie_top & "realiseret timer ialt"
				end if
				
				strJobLinie_top = strJobLinie_top & "</td>"
				
				
				if cint(visfakbare_res) = 1 then
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>Omsætning&nbsp;</td>"
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>Balance&nbsp;</td>"
				end if
				
				if cint(visfakbare_res) = 2 AND cint(forjob) = 0 then
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>Ressourcetimer tildelt.&nbsp;</td>"
				strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#d6dff5>Balance. (Reg. timer / Ressource timer)&nbsp;</td>"
				end if
				
				
				
				x = 0
				
				
				if cint(visJob) = 0 then
				
						for v = 0 to v - 1
						Redim preserve medabTottimer(v)
						Redim preserve restimerTot(v)
						Redim preserve omsTot(v)
						Redim preserve fakbareTimerTot(v)
						Redim preserve normTimerTot(v)
						medabTottimer(v) = 0
						restimerTot(v) = 0
						omsTot(v) = 0
						fakbareTimerTot(v) = 0
						normTimerTot(v) = 0
						
						
						'** Medarb 1 Row ***
						strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms1&" bgcolor=#eff3ff><font color=red>"&medarbnavnognr(v)&"</font>&nbsp;"
						
						'if cint(visfakbare_res) = 1 then
						'strJobLinie_top = strJobLinie_top & "realiseret fak.bare timer"
						'else
						'strJobLinie_top = strJobLinie_top & "realiseret timer ialt"
						'end if
						
						strJobLinie_top = strJobLinie_top & "</td>"
						
						
						
						if cint(visfakbare_res) = 1 then
						strJobLinie_top = strJobLinie_top & "<td align=right "&tdstyleTimOms&" bgcolor=#ffffff>Omsætning&nbsp;</td>"
						end if
						
						
						if cint(visfakbare_res) = 2 AND cint(forjob) = 0 then
						strJobLinie_top = strJobLinie_top & "<td class=lille align=right "&tdstyleTimOms&" bgcolor=#ffff99>Ressource timer tildelt</td>"
						strJobLinie_top = strJobLinie_top & "<td class=lille align=right "&tdstyleTimOms&" bgcolor=#ffff99>Balance reg.timer/ressource timer</td>"
						
						strJobLinie_top = strJobLinie_top & "<td class=lille align=right "&tdstyleTimOms&" bgcolor=Honeydew>Normeret timer ialt i periode</td>"
						strJobLinie_top = strJobLinie_top & "<td class=lille align=right "&tdstyleTimOms&" bgcolor=Honeydew>Balance (Reg. timer / Ressource timer)</td>"
						
						
						
									'**** Normeret timer ****
									sumnomtimerThis = 0
									strSQL3 = "SELECT normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, "_
									&" normtimer_lor, normtimer_son FROM medarbejdertyper WHERE id = " & jobmedtimer(x,15)
									
									'Response.write strSQL3
									'Response.flush
									oRec3.open strSQL3, oConn, 3 
									
									if not oRec3.EOF then
										sumnomtimerThis = ((oRec3("normtimer_man") * antalMandage) + (oRec3("normtimer_tir") * antalTirsdage ) + (oRec3("normtimer_ons") * antalOnsdage) + (oRec3("normtimer_tor") * antalTorsdage) + (oRec3("normtimer_fre") * antalFredage) + (oRec3("normtimer_lor") * antalLordage) + (oRec3("normtimer_son") * antalSondage)) 
									end if
									oRec3.close 
									
									normTimerTot(v) = sumnomtimerThis
									'sumnomtimerBal = (sumnomtimerThis - jobtimerIalt)
									
						
						end if
						
						next
						
				end if%>
				
				
				
				
				
				
				<%
				'************************************************************
				'*** Udskriver en ny linie med de valgte job / Aktiviteter
				'************************************************************ 
				'j = 0
				p = 0
				lastV = -1
				lastjid = 0
				
				for x = 0 to UBOUND(jobmedtimer, 1)
				
				
					
					
				
								if cint(forjob) = 0 then '** Job
								jobaktId = jobmedtimer(x,0)
								jobaktNr = jobmedtimer(x,6)
								jobaktNavn = jobmedtimer(x,1)
								jobaktType = jobmedtimer(x, 8)
								
								
								else '*** Udspeciffer Akt.
								
								jobaktId = jobmedtimer(x,12)
								jobaktNr = jobmedtimer(x,12)
								jobaktNavn = jobmedtimer(x,13)
								jobaktType = jobmedtimer(x, 14)
								
								end if
									
									
									if lastjid <> jobaktId AND len(jobaktNavn) <> 0  then
									lastjidTimerX = jobmedtimer(x,16)
									end if
						
									
									if (lastjidTimerX > 0 AND visnulfilter = 0) OR visnulfilter = 1 then 
									
									
									if lastjid <> jobaktId AND len(jobaktNavn) <> 0  then
									
									strJobLinie = strJobLinie & "</tr><tr>"
									strJobLinie = strJobLinie & "<td style='padding:2px; border-top:1px #999999 solid;' class=lille bgcolor='"& bgtd &"'>"&jobaktNavn&"&nbsp;("&jobaktNr&") "& jobaktType &"</font><font class=megetlillesilver><br>"& jobmedtimer(x, 20) &"</font></td>"
									
									
									strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#ffffff'>"& formatnumber(jobmedtimer(x, 19),2) &"</td>"
									
									if cint(visfakbare_res) = 1 then
									strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#ffffff'>"& formatnumber(jobmedtimer(x, 18),2) &"</td>"
									end if
									
									totbudget = totbudget + jobmedtimer(x, 18)
									totbudgettimer = totbudgettimer + jobmedtimer(x, 19)
										
											
											'*** Timer ***
											jobtimerIalt = 0
											jobOmsIalt = 0
											
											totaltotaljboTimerIalt = totaltotaljboTimerIalt + jobmedtimer(x, 16) ' + jobtimerIalt
											
											strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#eff3ff'>"& formatnumber(jobmedtimer(x, 16), 2) &"</td>"
											
											
											
											if cint(visfakbare_res) = 1 then
											
											'*** Omsætning  **
											totaltotaljobOmsIalt = totaltotaljobOmsIalt + jobmedtimer(x, 17) '+ jobOmsIalt
											strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#ffffff'>"&formatnumber(jobmedtimer(x, 17), 2)&"</td>"
											
											
											'*** Balance ***
											dbal = (jobmedtimer(x, 17) - jobmedtimer(x, 18))
											strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#ffdfdf>"& formatnumber(dbal, 2) &"</td>"
											
											dbialt = dbialt + (dbal)
											
											
											end if
											
											
											
											'*** Ressource timer vises kun på job niveau **
											if cint(visfakbare_res) = 2 AND cint(forjob) = 0  then
											
											strSQL3 = "SELECT sum(r.timer) AS restimer"_
											&" FROM ressourcer r WHERE r.jobid = "& jobmedtimer(x,0) &" "_
											&" AND r.dato BETWEEN '"& stdato &"' AND '"& sldato &"'"_
											&" GROUP BY r.jobid"
									
											'Response.write strSQL3
											'Response.flush
											
											jobRestimerThis = 0
											oRec3.open strSQL3, oConn, 3 
											if not oRec3.EOF then 
											jobRestimerThis = oRec3("restimer")
											end if
											oRec3.close
											
											totaltotalJobRestimer = totaltotalJobRestimer + jobRestimerThis
											balJobRestimer = (jobRestimerThis - jobtimerIalt)
											
													
											
											strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#ffff99' width=60>"&formatnumber(jobRestimerThis, 2)&"</td>"
											strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#ffff99' width=60>"&formatnumber(balJobRestimer, 2)&"</td>"
											
											end if
											
											
											lastjid = jobaktId
										end if
										
								
								
								
				
						'*************************************************************
						'**** Medarbjedere LOOP
						'**************************************************************
						
						if cint(visJob) = 0 then
						
							for v = 0 to v - 1
								
								
								if medarb(v) = jobmedtimer(x,4) then
								'*** timer total ***
								
								strJobLinie = strJobLinie & "<td class=lille align=right  "& tdstyleTimOms3 &" bgcolor='#eff3ff'>"
								
								if jobmedtimer(x,3) > 0 then
									strJobLinie = strJobLinie & formatnumber(jobmedtimer(x,3), 2)
								else
									strJobLinie = strJobLinie & "&nbsp;"
								end if 
								
								strJobLinie = strJobLinie & "</td>"
								
								 
									'*** fakbare timer ***
									if cint(visfakbare_res) = 1 then
									
									
									'**** Beløb / Omsætning ****
									strJobLinie = strJobLinie &"<td class=lille align=right "& tdstyleTimOms2 &" bgcolor='#ffffff'>"
									
									if jobmedtimer(x,7) > 0 then
									strJobLinie = strJobLinie & formatnumber(jobmedtimer(x,7), 2)
									else
									strJobLinie = strJobLinie & "&nbsp;"
									end if 
									strJobLinie = strJobLinie & "</td>"
								
								end if%>
								
								
								
								<%if cint(visfakbare_res) = 2 AND cint(forjob) = 0  then
									
									strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms2 &" bgcolor='#ffff99'>"
									
									'**** Ressource timer ****
									call restimer(jobmedtimer(x, 4), jobmedtimer(x, 0), sqlDatoStart, sqlDatoSlut)
									
									strJobLinie = strJobLinie & "</td>"
									
									balThis = (restimerThis - jobmedtimer(x,3))
									if len(balThis) <> 0 then
									balThis = balThis
									else
									balThis = 0
									end if
									
									strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms2 &" bgcolor='#ffff99'>"
									if balThis <> 0 then
									strJobLinie = strJobLinie & formatnumber(balThis, 2)
									else
									strJobLinie = strJobLinie & "&nbsp;"
									end if
									
									strJobLinie = strJobLinie & "</td>"
									
									
									strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms2 &" bgcolor='Honeydew'>&nbsp;-</td>"
									strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms2 &" bgcolor='Honeydew'>&nbsp;-</td>"
											
									
								end if
								
								
								medabTottimer(v) = medabTottimer(v) + jobmedtimer(x,3)
								restimerTot(v) = restimerTot(v) + restimerThis
								omsTot(v) = omsTot(v) + jobmedtimer(x,7)
								
								'fakbareTimerTot(v) = fakbareTimerTot(v) + jobmedtimer(x,9)
								
								
								'*** For at krydstjekke bliver restimer og balance summeret som totaler på medarb
								'*** Mens Timer ialt og fakbare timer og Oms. bliver summeret på jobtotaler
								restimerTotTot = restimerTotTot + restimerThis
								end if
								
							
							next 'medarb
						end if 'if visMedarb = 1 then
						
					response.flush
						
						end if 'jobmedtimer(x,16) og Nulfilter
				
				'*** Tilprint ***
				next 'job
			
			strJobLinie = strJobLinie & "</tr>"
			
			
			'***************************************
			'**** Udskriver Tables *****************
			'***************************************
			Response.write strJobLinie_top
			
			'*** Skal job vises eller vises kun totaler?
			 Response.write strJobLinie
			
			'***************************
			'*** totaler i bunden ******
			'***************************
			strJobLinie_total = "<tr>"
				strJobLinie_total = strJobLinie_total & "<td style='padding:2px; border-top:1px #999999 solid;' valign=bottom bgcolor=#d6dff5><b>Total:</b></td>"
						
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"& formatnumber(totbudgettimer, 2)&"</td>"
						
						if cint(visfakbare_res) = 1 then
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"& formatnumber(totbudget, 2)&"</td>"
						end if
						
						
						'*** Jobtotaler IALT ***
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"& formatnumber(totaltotaljboTimerIalt, 2)&"</td>"
						
						
						if cint(visfakbare_res) = 1 then
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"&formatnumber(totaltotaljobOmsIalt, 2)&"</td>"
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"&formatnumber(dbialt, 2)&"</td>"
						end if
						
						if cint(visfakbare_res) = 2 AND cint(forjob) = 0  then
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"&formatnumber(restimerTotTot, 2)&"</td>"
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#d6dff5>"
						
						balanceTotTot = (restimerTotTot - totaltotaljboTimerIalt)
						
						strJobLinie_total = strJobLinie_total & formatnumber(balanceTotTot, 2) & "</td>"
						
						end if
						
				
				
				
				if cint(visJob) = 0 then
				
						for v = 0 to v - 1
						strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms3&" bgcolor=#eff3ff>"&formatnumber(medabTottimer(v), 2)&"</td>"
						
							
							
							if cint(visfakbare_res) = 1 then
							strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#ffffff>"&formatnumber(omsTot(v), 2)&"</td>"
							end if
						
							if cint(visfakbare_res) = 2 AND cint(forjob) = 0  then
								
								'*** Ressource timer ****
								strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#ffff99>"&formatnumber(restimerTot(v), 2)&"</td>"
								
								balThisTot = (restimerTot(v) - medabTottimer(v))
								
								if len(balThisTot) <> 0 then
								balThisTot = balThisTot
								else
								balThisTot = 0
								end if
								
								strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=#ffff99>"&formatnumber(balThisTot, 2)&"</td>"
								
								'****  Normeret timer ***** 
								strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=Honeydew>"&formatnumber(normTimerTot(v), 2)&"</td>"
								
								NormBalThisTot = (normTimerTot(v) - medabTottimer(v))
								if len(NormBalThisTot) <> 0 then
								NormBalThisTot = NormBalThisTot
								else
								NormBalThisTot = 0
								end if
								
								strJobLinie_total = strJobLinie_total & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor=Honeydew>"&formatnumber(NormBalThisTot, 2)&"</td>"
							end if
						next
				  end if 'if visMedarb = 1 then
			strJobLinie_total = strJobLinie_total & "</tr></table></td></tr></table>"
			
			
			Response.write strJobLinie_total
			
			
			
			%>
			<form action="printversion.asp" method="post" name=theForm onsubmit="BreakItUp()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			<input type="hidden" name="txt1" id="txt1" value="<%=strJobLinie_top%>">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=strJobLinie%>">
			<input type="hidden" name="txt20" id="txt20" value="<%=strJobLinie_total%>">
			
			<input type="submit" value="Print version">
			
			</form>
			
			<b>Jobtyper:</b> Fastpris eller vægtet timepris ved løbende timer.<br><br>&nbsp;
			</div>
		<%end if ' max size%>	
		
	
		
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->