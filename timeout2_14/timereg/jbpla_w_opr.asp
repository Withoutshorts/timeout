<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="exch_meeting_webdav.asp"-->
<!--#include file="exch_meeting_webdav_del.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	select case func
	case "oprdb", "reddb"
	
	sttid = request("FM_sttid")
	sltid = request("FM_sltid")
	
	jobid = request("FM_jobid")
	medid = request("FM_mid")
	dato = request("FM_dato")
	aktid = request("FM_aktid")
	
	'*** jobnavn ***
	strSQLJ = "SELECT jobnavn, jobnr, k.kkundenavn, k.kkundenr, "_
	&" k.adresse, k.postnr, k.city, j.jobknr, j.beskrivelse, j.kundekpers, "_
	&" kp.navn, kp.email, kp.dirtlf, kp.mobiltlf FROM job j "_
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	&" LEFT JOIN kontaktpers kp ON (kp.id = j.kundekpers) WHERE j.id = "& jobid 
	'Response.write strSQLJ
	'Response.flush
	oRec.open strSQLJ, oConn, 3 
	if not oRec.EOF then
		
		jobnavnognr = oRec("jobnavn") & " ("& oRec("jobnr") &")"
		kundenavnognr = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"
		jobbesk = oRec("beskrivelse")
		kundenavnogadr = oRec("kkundenavn") & ", " & oRec("adresse") &", "& oRec("postnr") &" "& oRec("city")
		kundeadr = oRec("adresse") &", "& oRec("postnr") &" "& oRec("city")
		
		kpers = oRec("navn")
		kpers_email = oRec("email")
		kpers_dirtlf = oRec("dirtlf")
		kpers_mobiltlf = oRec("mobiltlf")
		
	end if
	oRec.close
	
	'*** Aktnavn ***
	strSQLA = "SELECT navn FROM aktiviteter WHERE id = "& aktid
	oRec.open strSQLA, oConn, 3 
	if not oRec.EOF then
	aktnavn = oRec("navn") 
	end if
	oRec.close
	
	
	timerthis_temp = datediff("n", sttid, sltid)
	timerthis_temp = (timerthis_temp/60)
	
	timerthis = replace(timerthis_temp, ",",".")
	
	sttid = formatdatetime(sttid, 3)
	sltid = formatdatetime(sltid, 3)
	
	if len(request("FM_gentag")) <> 0 then
	gentag = request("FM_gentag")
	else
	gentag = 0
	end if
	
	antal = request("FM_antal")
	interval = request("FM_interval")
	
	
	'*** Afsender ***
	if lto = "kringit" then
		seEmail = "connect@kringitsolutions.dk"
		seNavn = "KRING Timeout"
	else
		strSQL = "SELECT mid, mnavn, email FROM medarbejdere WHERE mid = "& session("mid")
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		seEmail = oRec("email")
		seNavn = oRec("mnavn")
		end if
		oRec.close
	end if
	
	'*** Modtager ****
	strSQL = "SELECT mid, mnavn, email, exchkonto FROM medarbejdere WHERE mid = "& medid
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	rcEmail = oRec("email")
	rcNavn = oRec("mnavn")
	rcDomKonto = replace(oRec("exchkonto"), "#", "\") '*** Til Exchange integration
	end if
	oRec.close
			
	exch_appid = "timeout_"& datepart("d", now) & "-" & datepart("m", now) & "-" & datepart("yyyy", now) & "_kl_" & datepart("h", now) & "_" & datepart("n", now) & "_" & datepart("s", now) 
		
	
	if func = "reddb" then
		
		emailsubtxt = "Booking (opdateret)"
		
		'*** Finder datoforskel ***
		strSQL5 = "SELECT dato, starttp, sluttp, exch_appid FROM ressourcer WHERE id ="& id
		oRec.open strSQL5, oConn, 3 
		while not oRec.EOF 
		
		datoDifference = dateDiff("d", oRec("dato"), dato)
		klokDiff_start = dateDiff("n", oRec("starttp"), sttid)
		klokDiff_slut = dateDiff("n", oRec("sluttp"), sltid)
		exch_appid = oRec("exch_appid")
		
		oRec.movenext
		wend
		oRec.close 
		
			
			'*** Opdaterer gentagelser ***
			if gentag <> 1 then
					
					
					strSQL = "SELECT r.exch_appid FROM ressourcer r WHERE id = "&id
					oRec.open strSQL, oConn, 3 
					if not oRec.EOF then
						exch_appid_opr = oRec("exch_appid")
					end if
					oRec.close 
					
			
					'** Opdater basis bookning (den aktuelle bookning) ****
					strSQL = "UPDATE ressourcer SET exch_appid = '"& exch_appid &"', jobid = "& jobid &", aktid = "& aktid &", timer = "& timerthis &", mid = "& medid &", starttp = '"& sttid &"', sluttp = '"& sltid &"', dato = '"& dato &"', serieid = 0 WHERE id ="&id
					oConn.execute(strSQL)
					
					
					'*** Sletter den gamle på Exchange **
					if len(exch_appid_opr) <> 0 then
						call delExchange(rcDomKonto, exch_appid_opr)
					end if
					
					'*** Exchange integration ***
					'**** Opretter mødet på Exchangeserver ****
					'Response.write rcDomKonto&","& exch_appid &", "& dato &", "& sttid & ", "& sltid
					'Response.flush
					if request("FM_exchange_nogo") <> "3" then
						call oprExchange(rcDomKonto, exch_appid, dato, sttid, sltid)
					end if
					'**********
			
			else
					
					serieid = request("serieid")
					
					strSQL5 = "SELECT id, dato, starttp, sluttp FROM ressourcer WHERE id = " & serieid & " OR (serieId = " & serieid & " AND serieId <> 0)" 
					oRec.open strSQL5, oConn, 3 
					x = 1
					while not oRec.EOF 
					
					nyDatoTemp = dateAdd("d", datoDifference, oRec("dato"))
					nyDato = year(nyDatoTemp) &"/"& month(nyDatoTemp) &"/" & day(nyDatoTemp)
					
					nyStartTid = dateAdd("n", klokDiff_start, oRec("starttp"))
					nySlutTid = dateAdd("n", klokDiff_slut, oRec("sluttp"))
						
						strSQL2 = "SELECT r.exch_appid FROM ressourcer r WHERE id = "& oRec("id")
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
							exch_appid_opr = oRec2("exch_appid")
						end if
						oRec2.close 
						
						strSQL = "UPDATE ressourcer SET exch_appid = '"& exch_appid_opr &"',"_
						&" jobid = "& jobid &", aktid = "& aktid &", "_
						&" timer = "& timerthis &", mid = "& medid &", starttp = '"& nyStartTid &"', "_
						&" sluttp = '"& nySlutTid &"', dato = '"& nyDato &"' WHERE id ="& oRec("id")
						oConn.execute(strSQL)
						
						'*** Sletter den gamle på Exchange **
						if len(exch_appid_opr) <> 0 then
							call delExchange(rcDomKonto, exch_appid_opr)
						end if
						
						'*** Exchange integration ***
						'**** Opretter mødet på Exchangeserver ****
						exch_appid_serie = exch_appid &"-"&x
						if request("FM_exchange_nogo") <> "3" then
							call oprExchange(rcDomKonto, exch_appid_serie, nyDato, nyStartTid, nySlutTid)
						end if
						'**********
					
					x = x + 1	
					oRec.movenext
					wend
					oRec.close 
				
					
			end if
			
			
			
	else
		
		'exch_appid = "timeout_"& datepart("d", now) & "-" & datepart("m", now) & "-" & datepart("yyyy", now) & "_kl_" & datepart("h", now) & "_" & datepart("n", now) & "_" & datepart("s", now) 
		
		emailsubtxt = "Booking"
		
		strSQL = "INSERT INTO ressourcer "_
		&" SET jobid = "& jobid &", aktid = "& aktid &", timer = "& timerthis &", "_
		&" mid = "& medid &", starttp = '"& sttid &"', "_
		&" sluttp = '"& sltid &"', dato = '"& dato &"', exch_appid = '"& exch_appid &"'"
		oConn.execute(strSQL)
			
			serieID = 0
			strSQL2 = "SELECT id FROM ressourcer WHERE id <> 0 ORDER BY id DESC"
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
				serieID = oRec2("id")
			end if
			oRec2.close 
			
			
			'*** Exchange integration ***
			'**** Opretter, redigerer mødet på Exchangeserver ****
			if request("FM_exchange_nogo") <> "3" then
				call oprExchange(rcDomKonto, exch_appid, dato, sttid, sltid)
			end if
			'**********
			
			'*** Gentagelser ****
			exch_appid_opr = exch_appid
			nyDatoTemp = dato
			if gentag = "1" then
				for x = 1 to antal - 1
					
					nyDato = nyDatoTemp
					
					select case interval
					case "1"
					nyDatoTemp = dateadd("ww", 1, nyDato)
					case "2"
					nyDatoTemp = dateadd("m", 1, nyDato)
					case "3"
					nyDatoTemp = dateadd("yyyy", 1, nyDato)
					case "4"
					nyDatoTemp = dateadd("d", 1, nyDato)
					case "5"
					nyDatoTemp = dateadd("q", 1, nyDato)
					end select
					
					nyDato = year(nyDatoTemp) &"/"& month(nyDatoTemp) &"/" & day(nyDatoTemp)
					exch_appid_serie = exch_appid_opr &"-"&x
					
						strSQL4 = "INSERT INTO ressourcer SET jobid = "& jobid &", aktid = "& aktid &", "_
						&" timer = "& timerthis &", mid = "& medid &", "_
						&" starttp = '"& sttid &"', sluttp = '"& sltid &"', dato = '"& nyDato &"', serieid = "& serieID &", exch_appid = '"& exch_appid_serie &"'"
						oConn.execute(strSQL4)
						
						if request("FM_exchange_nogo") <> "3" then
							call oprExchange(rcDomKonto, exch_appid_serie, nyDato, sttid, sltid)
						end if
				
				next
			end if
	
	end if
	
	Response.flush
	
	
	
	
	
	'*** Emnail notifikation ***
	if request("FM_email_nogo") <>  "3" then
	
	
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					
	' Sætter Charsettet til ISO-8859-1 
	Mailer.CharSet = 2
	' Afsenderens navn 
	Mailer.FromName = ""& seNavn
	' Afsenderens e-mail 
	Mailer.FromAddress = ""& seEmail
	Mailer.RemoteHost = "webmail.abusiness.dk"
	' Modtagerens navn og e-mail
	Mailer.AddRecipient rcNavn, rcEmail
	'Mailer.AddBCC "Support", "support@outzource.dk" 
	' Mailens emne
	Mailer.Subject = ""& emailsubtxt &": d. " & formatdatetime(dato, 2) & " kl. " & sttid & " - "& sltid
	
	Mailer.ContentType = "text/html"
	
	' Selve teksten
	Mailer.BodyText = "<html><head><title></title></head><body>" & "Hej "& rcNavn & "<br>"_ 
	& "Du er blevet booket på følgende job.<br><br>"_
	& "Kunde: " & kundenavnognr &"<br>"_ 
	& "Adr: " & kundeadr &"<br>"_ 
	& "Kontakt: " & kpers &"<br>"_  
	& "Email: " & kpers_email &"<br>"_ 
	& "Dir. Tlf: "& kpers_dirtlf &"<br>"_ 
	& "Mobil Tlf: "& kpers_mobiltlf & "<br><br>"_ 
	& "Job: <b>"& jobnavnognr &"</b><br>"_ 
	& "Beskrivelse:<br> "& jobbesk &"<br><br>"_
	& "Dato: d. " & formatdatetime(dato, 2) & " kl. " & sttid & " - "& sltid  & "<br><br>"_ 
	
	& "Link til timeregistrering: <br>"_ 
	& "https://outzource.dk/"& lto & "<br><br>"_ 
	
	& "Med venlig hilsen"& "<br><br>" & seNavn & "<br><br>&nbsp;" _
	& "</body></html>" 
	
	'& "Link til kalender (rediger denne bookning): " & vbCrLf _
	'& "https://outzource.dk/"& lto & "" & vbCrLf & vbCrLf _ 
	
	Mailer.SendMail
	
	end if 'email nogo
	
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
   	
	
   
	
	case "step3"
	
	
	sttid = request("klok_timer")&":"&request("klok_min")&":00"
	sltid = request("klok_timer_slut")&":"&request("klok_min_slut")&":00"
	
	if formatdatetime(sltid, 3) <= formatdatetime(sttid, 3) then
	sltid = dateadd("n", 15, dato&" "&sttid)
	else
	sltid = sltid
	end if
	
	gentag = request("FM_gentag")
	antal = request("FM_antal")
	interval = request("FM_interval")
	serieid = request("serieid")
	
	jobid = request("FM_opr_jobid") 'request("FM_jobid")
	medid = request("FM_opr_mid")
	dato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
	'dato = request("FM_dato")
	'sttid = request("FM_sttid")
	'sltid = request("FM_sltid")
	'aktid = request("FM_aktid")
	
	if id <> 0 then
	func = "reddb"
	funcpil = "opdaterpil"
	else
	func = "oprdb"
	funcpil = "opretpil"
	end if
	
	gentag = request("FM_gentag")
	antal = request("FM_antal")
	interval = request("FM_interval")
	serieid = request("serieid")
	
	stthis = request("step")
	
	antalstepialt = request("stepialt")
	%>
	
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<font class="stor-blaa">Step <%=stthis%>/<%=antalstepialt%></font>
	<div name="step3" id="step3" style="position:absolute; display:; visibility:visible; left:20px; top:40px; background-color:#eff3ff; width:600px; z-index:10000; border:1px #5582d2 solid; padding:20px;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="300">
	<form action="jbpla_w_opr.asp?func=<%=func%>" method="post" name="opr" id="opr">
	<input type="hidden" name="id" id="id" value="<%=id%>">
	<input type="hidden" name="FM_dato" id="FM_dato" value="<%=dato%>">
	<input type="hidden" name="FM_sttid" id="FM_sttid" value="<%=sttid%>">
	<input type="hidden" name="FM_sltid" id="FM_sltid" value="<%=sltid%>">
	<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=jobid%>">
	<input type="hidden" name="FM_aktid" id="FM_aktid" value="0"> <!--<aktid%>-->
	
	<input type="hidden" name="FM_gentag" id="FM_gentag" value="<%=gentag%>">
	<input type="hidden" name="FM_antal" id="FM_antal" value="<%=antal%>">
	<input type="hidden" name="FM_interval" id="FM_interval" value="<%=interval%>">
	<input type="hidden" name="serieid" id="serieid" value="<%=serieid%>">
	
	
	<tr><td colspan=2><b>Vælg medarbejder:</b><br>
	Vælg den medarbejder der skal udføre jobbet.<br>
	Der kan vælges mellem alle de medarbejdere der har adgang til det valgte job.<br>&nbsp;</td></tr>
	<tr><td colspan=2>
	<%
			strSQL = "SELECT mnavn, mid, mnr, projektgrupper.id AS id, "_
			&" projektgrupper.navn AS navn, job.projektgruppe1, job.projektgruppe2, "_
			&" job.projektgruppe3, job.projektgruppe4, job.projektgruppe5, "_
			&" job.projektgruppe6, job.projektgruppe7, job.projektgruppe8, "_
			&" job.projektgruppe9, job.projektgruppe10 FROM job "_
			&" LEFT JOIN projektgrupper ON (projektgrupper.id = job.projektgruppe1"_
			&" OR projektgrupper.id = job.projektgruppe2 "_
			&" OR projektgrupper.id = job.projektgruppe3 OR projektgrupper.id = job.projektgruppe4 "_
			&" OR projektgrupper.id = job.projektgruppe5 OR projektgrupper.id = job.projektgruppe6 "_
			&" OR projektgrupper.id = job.projektgruppe7 OR projektgrupper.id = job.projektgruppe8 "_
			&" OR projektgrupper.id = job.projektgruppe9 OR projektgrupper.id = job.projektgruppe10) "_
			&" LEFT JOIN progrupperelationer ON (progrupperelationer.projektgruppeid = projektgrupper.id)"_
			&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid AND mansat <> 2) "_
			&" WHERE job.id = "& jobid
			
			'Response.write strSQL
			'Response.flush
			
			%>
			<select name="FM_mid" id="FM_mid">
			<%
			oRec.open strSQL, oConn, 3
			t = 0
			While not oRec.EOF 
			
			if len(oRec("mnavn")) <> 0 then
				if func = "oprdb" then
					if t = 0 then
					ts = "SELECTED"
					else
					ts = ""
					end if
				else
					if cint(medid) = oRec("mid") then
					ts = "SELECTED"
					else
					ts = ""
					end if
				end if
			
			%>
			<option value="<%=oRec("mid")%>" <%=ts%>><%=oRec("mnavn")%></option>
			<%
			t = t + 1
			end if
			oRec.movenext
			wend
			oRec.close 
	%>
	</select><br><br>
	<input type="checkbox" name="FM_email_nogo" id="FM_mail_nogo" value="3">Send <b>ikke</b> mail til medarbejder vedr. booking.<br>
	<input type="checkbox" name="FM_exchange_nogo" id="FM_exchange_nogo" value="3" CHECKED>Opret <b>ikke</b> denne booking i kalender på Exchange server.
	<br>&nbsp;	
	</td></tr>
	<tr>
		<%if t = 0 then%>
		<td colspan=2><br>
		<FONT color=red><b>Fejl!</b><br>Der er ikke fundet medarbejdere med adgang til denne aktivitet..</font>
		<%else%>
		<td colspan=2><br>
		<input type="image" src="../ill/<%=funcpil%>.gif">
		<%end if%>
		</td>
	</tr>
	</form>
	</table>
	
	</div>
	<div name="tb" id="tb" style="position:absolute; display:; visibility:visible; left:20px; top:320px; background-color:#ffffff;">
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><br>&nbsp;
	</div>
	
	<%
	'case "step2"
	
	'if len(request("aid")) <> 0 then
	'aid = request("aid")
	'else
	'aid = 0
	'end if
	
	'jobid = request("FM_opr_jobid")
	'medid = request("FM_opr_mid")
	'dato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
	
	'sttid = request("klok_timer")&":"&request("klok_min")&":00"
	'sltid = request("klok_timer_slut")&":"&request("klok_min_slut")&":00"
	
	'if formatdatetime(sltid, 3) <= formatdatetime(sttid, 3) then
	'sltid = dateadd("n", 15, dato&" "&sttid)
	'else
	'sltid = sltid
	'end if
	
	'gentag = request("FM_gentag")
	'antal = request("FM_antal")
	'interval = request("FM_interval")
	'serieid = request("serieid")
	
	'func = "step3"
	'stthis = request("step")

	%>
	<!--
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<font class="stor-blaa">Step <=stthis%> / 4(5)</font>
	<div name="step2" id="step2" style="position:absolute; display:; visibility:visible; left:20; top:40; background-color:#eff3ff; width:220; z-index:10000; border:1px #5582d2 solid; padding:5px;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="300">
	<form action="jbpla_w_opr.asp?func=<=func%>&step=<=stthis+1%>" method="post" name="opr" id="opr">
	<input type="hidden" name="id" id="id" value="<=id%>">
	<input type="hidden" name="FM_dato" id="FM_dato" value="<=dato%>">
	<input type="hidden" name="FM_sttid" id="FM_sttid" value="<=sttid%>">
	<input type="hidden" name="FM_sltid" id="FM_sltid" value="<=sltid%>">
	<input type="hidden" name="FM_opr_mid" id="FM_opr_mid" value="<=medid%>">
	<input type="hidden" name="FM_jobid" id="FM_jobid" value="<=jobid%>">
	
	<input type="hidden" name="FM_gentag" id="FM_gentag" value="<=gentag%>">
	<input type="hidden" name="FM_antal" id="FM_antal" value="<=antal%>">
	<input type="hidden" name="FM_interval" id="FM_interval" value="<=interval%>">
	<input type="hidden" name="serieid" id="serieid" value="<=serieid%>">
	
	<tr>
		<td colspan=2><b>Vælg den aktivitet der ønskes udført.</b><br>
		Der kan vælges en bestemt aktivitet som ønskes udført, eller der kan vælges en generel aktivitet.<br>&nbsp;</td>
	</tr>
	<tr>
		<td align=right style="padding-right:3px;">Aktivitet:</td>
		<td><select name="FM_aktid" id="FM_aktid" style="width:250;">
			<%
				'strSQLJ = "SELECT id, navn FROM aktiviteter WHERE job = "& jobid &" AND aktstatus = 1 AND fakturerbar <> 2 ORDER BY navn"
				'oRec.open strSQLJ, oConn, 3 
				'a = 0
				'while not oRec.EOF 
				'if cint(aid) = oRec("id") then
				'aidsel = "SELECTED"
				'else
				'aidsel = ""
				'end if%>
				<option value="<=oRec("id")%>" <=aidsel%>><=left(oRec("navn"), 30)%></option>
				<%
				'a = a + 1
				'oRec.movenext
				'wend
				'oRec.close 
				%>
			</select></td>
	</tr>
	<tr>
	<if a = 0 then%>
	<td colspan=2><br>
	<FONT color=red><b>Fejl!</b><br>Der er ikke oprettet aktiviteter på dette job.</font>
	<else%>
	<td colspan=2 align=center><br><input type="image" src="../ill/naestepil.gif">
	<end if%>
	</td></tr>
	</form>
	</table>
	
	</div>
	<div name="tb" id="tb" style="position:absolute; display:; visibility:visible; left:20; top:280; background-color:#ffffff;">
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><br>&nbsp;
	</div>
	-->
	
	<%
	case else
	%> 
	<script LANGUAGE="javascript">
		function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
	</script>
	<%
	if len(request("medid")) <> 0 then
	medid = request("medid")
	else
	medid = 0
	end if
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	if len(request("aid")) <> 0 then
	aid = request("aid")
	else
	aid = 0
	end if
	
	if len(request("serieid")) <> 0 then
	serieid = request("serieid")
	else
	serieid = 0
	end if
	
	dato = request("dato")
	strDag = day(dato)
	strMrd = month(dato)
	strAar = year(dato)
	
	'Response.write dato
	
	strKlokkeslet = request("sttid")
	
	if func = "red" then 
	strKlokkeslet_slut = request("sltid")
	else
	strKlokkeslet_slut = dateadd("n", 30, dato&" "&strKlokkeslet)
	end if
	
	
	formatstrKlokkeslet = FormatDateTime(strKlokkeslet,4)
	klok_timer = left(formatstrKlokkeslet,2)
	klok_min = right(formatstrKlokkeslet,2)
	
	formatstrKlokkeslet_slut = FormatDateTime(strKlokkeslet_slut,4)
	klok_timer_slut = left(formatstrKlokkeslet_slut,2)
	klok_min_slut = right(formatstrKlokkeslet_slut,2)
	
	jobstKri = request("datostkri") 
	jobslKri = request("datoslkri")

%>
<!--#include file="inc/dato2_b.asp"-->

<%if func = "opr" then
stthis = "1"
antalstepialt = 3
else
stthis = request("step")
antalstepialt = request("stepialt")
end if%>

<br>&nbsp;&nbsp;&nbsp;&nbsp;<font class="stor-blaa">Step <%=stthis%>/<%=antalstepialt%></font>
<div name="opretny" id="opretny" style="position:absolute; display:; visibility:visible; left:20px; top:40px; background-color:#eff3ff; width:600px; z-index:10000; border:1px #5582d2 solid; padding:20px;">
	
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="400">
	
	<%if func = "opr" then%>
	
	<form method=post action="jobs.asp?showaspopup=y&func=opret&id=0&int=1&rdir=jbpla_w&sttid=<%=strKlokkeslet%>&dato=<%=dato%>&datostkri=<%=jobstKri%>&datoslkri=<%=jobslKri%>&stepialt=<%=antalstepialt%>">
	<tr>
		<td style="white-space:nowrap; font-size:12px;">
		<b>A)</b> <b>Opret nyt job Ekspress</b><br>
		Kunde:<select name="FM_kunde" id="FM_kunde" style=" width:300px;">
			<%strSQL = "SELECT kkundenavn, kkundenr, kid FROM "_
			&" kunder WHERE ketype <> 'e' ORDER BY kkundenavn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			%>
			<option value="<%=oRec("kid")%>"><%=oRec("kkundenavn")%> (<%=oREc("kkundenr") %>) </option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			
		</select>&nbsp;<input type="submit" value="Opret nyt Job >> " style="">
		</td>
	</tr></form>
	
		
	<tr>
		<td style="white-space:nowrap; font-size:12px;"><br /><br /><b>B)</b> <a href="jbpla_w_opr.asp?func=step1&sttid=<%=strKlokkeslet%>&dato=<%=dato%>&datostkri=<%=jobstKri%>&datoslkri=<%=jobslKri%>&id=0&step=2&stepialt=<%=antalstepialt%>">Brug eksisterende job >></a></td>
	</tr>
	<form method=post action="job_kopier.asp?func=kopier&fm_kunde=0&filt=0&showaspopup=y&sttid=<%=strKlokkeslet%>&dato=<%=dato%>&datostkri=<%=jobstKri%>&datoslkri=<%=jobslKri%>">
	<tr>
		<td style="font-size:12px;">
		<%if request("jobkopieret") = "j" then%>
		<div name="jkop" id="jkop" style="position:relative; display:; visibility:visible; background-color:#fffff1; width:300px; padding:10px; border:1px #003399 dashed">
		<font color="#2c962d"><b>Du har netop kopieret et job!</b><br>Gå videre ved at vælg pkt. B.<br>
		<br />Vælg pkt. C hvis du ønsker at kopiere endnu et job.</font><br>&nbsp;
		</div><br>
		<%end if%>
		<br><br><br>
		<b>C)</b> <b>Kopier eksisterende job?</b><br>
		Du kan vælge at kopiere et eksisterende job.
		Når jobbet er kopieret, vil du retunere til denne side, 
		hvorefter det nye job kan vælges ved at følge pkt. B.<br>  
		
		<select name="id" id="id" style=" width:300px;">
			<%strSQL = "SELECT jobnavn, jobnr, id, jobknr, kkundenavn, kkundenr, kid FROM job "_
			&"LEFT JOIN kunder ON (kid = jobknr) WHERE jobstatus = 1 AND fakturerbart = 1 ORDER BY kkundenavn, jobnavn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			%>
			<option value="<%=oRec("id")%>">[<%=oRec("kkundenavn")%>] -- <%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) </option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			
		</select>&nbsp;
		<input type="submit" value="Kopier job >>" style="">
		<!--<input type="image" src="../ill/soeg-knap.gif">-->
		<br><br>&nbsp;</td>
	</tr></form>
	<%else%>
	
	
	<form action="jbpla_w_opr.asp?func=step3&step=<%=stthis+1%>&stepialt=<%=antalstepialt%>" method="post" name="opr" id="opr">
	<input type="hidden" name="id" id="id" value="<%=id%>">
	<input type="hidden" name="aid" id="aid" value="<%=aid%>">
	<input type="hidden" name="FM_opr_mid" id="FM_opr_mid" value="<%=medid%>">
	<input type="hidden" name="serieid" id="serieid" value="<%=serieid%>">
	
	<tr>
		<td colspan=2><b>Vælg Dato, klokkeslet og job.</b><br>&nbsp;</td>
	</tr>
	<tr>
		<td align=right style="padding-right:3px;">Dato:</td>
		<td><select name="FM_start_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>"><%=strAar%></option>
		<option value="2002">2002</option>
		<option value="2003">2003</option>
	   	<option value="2004">2004</option>
	   	<option value="2005">2005</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
	</tr>	
			
	
	<tr><td align=right style="padding-right:3px;">Start kl: </td>
		<td><select name="klok_timer">
		<option value="<%=klok_timer%>" SELECTED><%=klok_timer%></option>
		<option value="08">08</option>
	   	<option value="09">09</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<!--<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>-->
	  	</select>&nbsp;<b>:</b>
		
		<select name="klok_min">
		<%if func <> "opret" then%>
		<option value="<%=klok_min%>" selected><%=klok_min%></option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
	</tr>
	
	
	<tr><td align=right style="padding-right:3px;">Slut kl:</td>
		<td>
		<select name="klok_timer_slut">
		<option value="<%=klok_timer_slut%>" SELECTED><%=klok_timer_slut%></option>
		<option value="08">08</option>
	   	<option value="09">09</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<!--<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>-->
	  	</select>&nbsp;<b>:</b>
		
		<select name="klok_min_slut">
		<%if func <> "opret" then%>
		<option value="<%=klok_min_slut%>" selected><%=klok_min_slut%></option>
		<%else%>
		<option value="30" SELECTED>30</option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select></td>
	</tr>
	
	
	
	<%if id = 0 then%>
	
	<tr><td colspan=2><br><br>
		<input type="checkbox" name="FM_gentag" id="FM_gentag" value="1"><b>Gentag booking?</b></td>
	</tr>
	<tr><td align=right style="padding-right:3px;">Gentag:</td>
		<td><input type="text" name="FM_antal" id="FM_antal" size=4> (Antal aftaler/bookinger ialt)</td>
	</tr>
	<tr><td align=right style="padding-right:3px;">Interval:</td>
		<td><select name="FM_interval" id="FM_interval">
		<option value="1" SELECTED>Ugenligt</option>
		<option value="2">Månedligt</option>
		<option value="3">Årligt</option>
		<option value="4">Dagligt</option>
		<option value="5">Kvartårligt</option>
		</select></td>
	</tr>
	<%else%>
	<tr><td colspan=2><br>
		<%
		strSQL = "SELECT id, dato FROM ressourcer WHERE id = " & serieid & " OR (serieId = " & serieid & " AND serieId <> 0)" 
		oRec.open strSQL, oConn, 3 
		s = 0
		while not oRec.EOF 
		s = s + 1
		oRec.movenext
		wend
		oRec.close 
		
		if s > 1 then%>
		<b>Der er <%=s%> bookninger i denne serie.</b><br>
		<input type="checkbox" name="FM_gentag" id="FM_gentag" value="1" CHECKED> Rediger alle i denne serie.<br>
		<font class=megetlillesort>Når serie redigeres flyttes alle bookninger det antal timer/dage frem eller tilbage som den aktuelle bookning flyttes.<br><br>
		Hvis der kun redigeres <u>1</u> bookning fra en serie, og denne IKKE er stam-bookningen, vil den bookning der redigeres ikke længere være en del af 
		den oprindelige serie.</font>
		<%else%>
		<input type="hidden" name="FM_gentag" id="FM_gentag" value="0">
		<%end if%>
		</td>
	</tr>
	<%end if%>
	<tr><td colspan=2><br><br>
		<b>[kontakt] -- Job.</b></td>
	</tr>
	<tr><td colspan=2 style="padding-right:3px;"><select name="FM_opr_jobid" id="FM_opr_jobid" style=" width:350;">
	
	<%
	
			strSQLJ = "SELECT jobnavn, jobnr, id, jobknr, kkundenavn, kkundenr, kid FROM job "_
			&"LEFT JOIN kunder ON (kid = jobknr) WHERE jobstatus = 1 AND fakturerbart = 1 ORDER BY kkundenavn, jobnavn" 'AND (jobstartdato <= '"& jobslKri &"' AND jobslutdato >= '"& jobstKri &"') 
			oRec.open strSQLJ, oConn, 3 
			while not oRec.EOF 
				
				if cint(jobid) = cint(oRec("id")) then
				ts2 = "SELECTED"
				else
				ts2 = ""
				end if
			%>
			<option value="<%=oRec("id")%>" <%=ts2%>>[<%=oRec("kkundenavn")%>] -- <%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) </option>
			<%
			oRec.movenext
			wend
			oRec.close %>
	
	</select>
	</td></tr>
	<tr><td colspan=2 align=center><br><input type="image" src="../ill/naestepil.gif"></td></tr>
	</form>
	<%end if%>
	
	</table>

	
	<%
	end select
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
