<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->

<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
		
		
	print = request("print")
		
		
    if print <> "j" then%>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <!--#include file="../inc/regular/topmenu_inc.asp"-->
    <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	    <%call tsamainmenu(9)%>
	    </div>
	    <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	    <%
	    call erptopmenu()
	    %>
	    </div>

    <%else%>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
    <%end if
		
		
	public totfaktimer, totfakbelob, totmedtimer, totmedbel 
	sub medarbejderTD(lastMid, thismedtimer, thismedbelob, kundeKri2, strMedNavn)
				
				strMedTimerogFak = "<td width=80 bgcolor=#ffffff align=right class=lille style='border:1px #999999 solid;'>"&formatnumber(thismedtimer)&" t.</td>"
				strMedTimerogFak = strMedTimerogFak &"<td class=lille bgcolor=#ffffff width=80 align=right style='border:1px #999999 solid;'>"&formatcurrency(thismedbelob)&"</td>"
				
				strSQL2 = "SELECT sum(fm.fak) AS medfaktimer, sum(fm.beloeb) AS medfakbel FROM fakturaer f, fak_med_spec fm WHERE "_
				&" (fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' "& kundeKri2 &") AND (fm.fakid = f.fid AND fm.mid = "& lastMid &") GROUP BY fm.mid"
				
				'Response.write strSQL2 &"<br><br>"
				
				medFakTimerTot = 0
				medFakBelTot = 0
				
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				
				medFakTimerTot = oRec2("medfaktimer")
				medFakBelTot = oRec2("medfakbel")
				
				end if
				oRec2.close 
			
				
				strMedTimerogFak = strMedTimerogFak &"<td width=80 bgcolor=#ffffff class=lille align=right style='border:1px #999999 solid;'>"&formatnumber(medFakTimerTot)&" t.</td>"
				strMedTimerogFak = strMedTimerogFak & "<td class=lille bgcolor=#ffffff align=right width=80 style='border:1px #999999 solid;'>"&formatcurrency(medFakBelTot)&"</td></tr>"
				
				
				if medFakTimerTot <> 0 OR thismedtimer <> 0 then
				Response.write strMedNavn
				Response.write strMedTimerogFak
				
				strMedNavn = ""
				strMedTimerogFak = ""
				end if
				
				
				totmedtimer = totmedtimer + thismedtimer
				totmedbel = totmedbel + thismedbelob
				totfaktimer = totfaktimer + medFakTimerTot
				totfakbelob = totfakbelob + medFakBelTot 
	end sub
	
	function medtotaler(sallkid)
	%><br><br>
	<img src="../ill/ac0016-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Medarbejder totaler:</b><br>
			<table cellpadding=1 cellspacing=1 border=0>
			<tr>
				<td bgcolor="#ffffff" width=180 valign=bottom class=lille style="padding-left:2px; border:1px #999999 solid;">Navn</td>
				<td bgcolor="#ffffff" width=160 colspan=2 class=lille style="border:1px #999999 solid;">Registreret <br>(fakturerbare)</td>
				<td bgcolor="#ffffff" width=160 colspan=2 class=lille valign=bottom style="padding-left:2px; border:1px #999999 solid;">Faktureret</td>
			</tr>
			<%
			
			if sallkid = 0 then
			jtype = "LEFT"
			kundeKri1 = ""
			kundeKri2 = ""
			else
			jtype = "LEFT"
			kundeKri1 = "  t.tknr = " & sallkid & " AND "
			kundeKri2 = " AND fakadr = " & sallkid 
			end if
			
			
			
			strSQL = "SELECT medarbejdere.mid AS medid, mnr, mnavn, sum(t.timer) AS sumtimer, "_
			&" timepris, j.fastpris, budgettimer, ikkebudgettimer, jobtpris FROM medarbejdere "_
			&" "& jtype &" JOIN timer t ON ("& kundeKri1 &" tmnr = medarbejdere.mid AND tfaktim = 1 "_
			&" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"')"_
			&" LEFT JOIN job j ON (jobnr = tjobnr) GROUP BY medarbejdere.mid, timepris, jobnr ORDER BY mnavn"
			
			'Response.write strSQL
			'Response.flush
			
			lastMid = 0
			v = 0
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			
			
			if lastMid <> oRec("medid") then
			
				if v > 0 then
					call medarbejderTD(lastMid, thismedtimer, thismedbelob, kundeKri2, strMedNavn)
					
					thismedtimer = 0
					thismedbelob = 0
				end if
			
			strMedNavn = "<tr><td bgcolor=#ffffff width=140 class=lille style='padding-left:2px; border:1px #999999 solid;'>"&oRec("mnavn")&" ("&oRec("mnr")&")</td>"
			end if%>
			
			<%
			thismedtimer = thismedtimer + oRec("sumtimer")
			
			if len(thismedtimer) <> 0 then
			thismedtimer = thismedtimer
			else
			thismedtimer = 0
			end if
			
			'Response.write oRec("medid") & ": "& thismedtimer & "<br>"
			
			if oRec("fastpris") <> 1 then
			thismedbelob = thismedbelob + (oRec("sumtimer") * oRec("timepris"))
			else
			thismedbelob = thismedbelob + (oRec("sumtimer") * (oRec("jobtpris")/oRec("budgettimer")))
			end if
			
			if len(thismedbelob) <> 0 then
			thismedbelob = thismedbelob
			else
			thismedbelob = 0
			end if
			
			lastMid = oRec("medid")
			
			v = 1
			oRec.movenext
			wend
			oRec.close 
			
			if v > 0 then
				call medarbejderTD(lastMid, thismedtimer, thismedbelob, kundeKri2, strMedNavn)
			end if
			
			if v = 0 then
			%>
			<tr>
				<td colspan=5 class=lille bgcolor="#ffffff"><font color="red">Der er ikke udsendt fakturaer i den valgte periode.</td>
			</tr>
			<%
			end if
			%>
			
			<tr>	
				<td bgcolor="#ffffff" class=lille style="border:1px #999999 solid;">Total:</td>
				<td bgcolor="#ffffe1" align=right class=lille style="border:1px #999999 solid;"><%=formatnumber(totmedtimer)%> t.</td>
				<td bgcolor="#ffffe1" align=right class=lille style="border:1px #999999 solid;"><%=formatcurrency(totmedbel)%></td>
				<td bgcolor="#ffffe1" align=right class=lille style="border:1px #999999 solid;"><%=formatnumber(totfaktimer)%> t.</td>
				<td bgcolor="#ffffe1" align=right class=lille style="border:1px #999999 solid;"><%=formatcurrency(totfakbelob)%></td>
			</tr>
	</table>
		
		 <br><br><br>&nbsp;
			
			
			
		
		<%
	end function	
		
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	 
	
	
	
	
	'*****************************************************************************
	'** Periode afgrænsning ***
	'*****************************************************************************
	 
	'*** Periode vælger er brugt ****
	 if len(request("FM_start_dag")) <> 0 then
	'* Bruges i weekselector *
	strMrd =  request("FM_start_mrd")
	strDag  = request("FM_start_dag")
	if len(request("FM_start_aar")) > 2 then
	strAar = right(request("FM_start_aar"), 2)
	else
	strAar = request("FM_start_aar")
	end if
	
	strMrd_slut =  request("FM_slut_mrd")
	strDag_slut  =  request("FM_slut_dag")
	
	if len(request("FM_slut_aar")) > 2 then
	strAar_slut = right(request("FM_slut_aar"),2)
	else
	strAar_slut = request("FM_slut_aar")
	end if
	
	
	select case strMrd
	case 2
		if strDag > 28 then
		strDag = 28
		else
		strDag = strDag
		end if
	case 4, 6, 9, 11
		if strDag > 30 then
		strDag = 30
		else
		strDag = strDag
		end if
	end select
	
	select case strMrd_slut
	case 2
		if strDag_slut > 28 then
		strDag_slut = 28
		else
		strDag_slut = strDag_slut
		end if
	case 4, 6, 9, 11
		if strDag_slut > 30 then
		strDag_slut = 30
		else
		strDag_slut = strDag_slut
		end if
	end select
	
	'*** Den valgte start og slut dato ***
	StrTdato = strDag &"/" & strMrd & "/20" & strAar
	StrUdato = strDag_slut &"/" & strMrd_slut & "/20" & strAar_slut
	
	else
	
		'*Brug cookie eller dagsdato?
		if len(Request.Cookies("datoer")("st_dag")) <> 0 then
	
		strMrd = Request.Cookies("datoer")("st_md")
		strDag = Request.Cookies("datoer")("st_dag")
		strAar = Request.Cookies("datoer")("st_aar") 
		strDag_slut = Request.Cookies("datoer")("sl_dag")
		strMrd_slut = Request.Cookies("datoer")("sl_md")
		strAar_slut = Request.Cookies("datoer")("sl_aar")
		
		else
		
		
			StrTdato = date-31
			StrUdato = date 
			
			'* Bruges i weekselector *
			if month(now()) = 1 then
			strMrd = 12
			else
			strMrd = month(now()) - 1
			end if
			
			strDag = day(now())
			
			if month(now()) = 1 then
			strAar = right(year(now()) - 1, 2)
			else
			strAar = right(year(now()), 2) 
			end if
			
			strMrd_slut = month(now())
			strAar_slut = right(year(now()), 2) 
			
			if strDag > "28" then
			strDag_slut = "1"
			strMrd_slut = strMrd_slut + 1
			else
			strDag_slut = day(now())
			end if
			
			
		end if
	end if
	
	'** Indsætter cookie **
	Response.Cookies("datoer")("st_dag") = strDag
	Response.Cookies("datoer")("st_md") = strMrd
	Response.Cookies("datoer")("st_aar") = strAar
	Response.Cookies("datoer")("sl_dag") = strDag_slut
	Response.Cookies("datoer")("sl_md") = strMrd_slut
	Response.Cookies("datoer")("sl_aar") = strAar_slut
	Response.Cookies("datoer").Expires = date + 10		
	
	
	'**** SQL datoer ***
	sqlSTdato = "20" & strAar &"/"& strMrd &"/"& strDag 
	sqlSLUTdato = "20" & strAar_slut &"/"& strMrd_slut &"/"& strDag_slut 
	
	
	if request("print") <> "j" then
	dleft = "20"
	dtop = "132"
	
	else
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="650">
				<tr>
					<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
					<td align=right style="padding-right:30px; padding-top:3px;">
					<a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a></td>
				</tr>
	</table>
	<%
	dleft = "10"
	dtop = "60"
	end if
	%>
	
		
		<div style="position:absolute; left:<%=dleft%>; top:<%=dtop%>;">
		<%if request("print") <> "j" then%>
		<img src="../ill/header_afst.gif" width="600" height="49" alt="" border="0"><br>
		<%else%>
		<h3>Afstemningsoversigt</h3>
		<a href="javascript:history.back()"><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0"></a>&nbsp;
		<%end if%>
		
		Oversigt over timeregistreringer og fakturerede timer for hver medarbejder i en given periode.<br>
		Afstemningsoversigten dækker kun oprindelige fakturaer, <b>IKKE eventuelle rykkere eller kreditnotaer.</b><br>
		
		
		
		<br>
<%	
	'*** Valgt kunde ***
	if len(request("FM_kunder")) <> 0 AND request("FM_kunder") <> 0 then
	selKunde = request("FM_kunder")
	selKundeKri = " kunder.kid = " & selKunde &""
	showall = 0
	else
	selKunde = 0
	'selKundeKri = ""
	showall = 1
	end if

if request("print") <> "j" AND showall = 1 then%>

<table cellspacing="0" cellpadding="0" border="0" bgcolor="#eff3ff">
<tr bgcolor="#5582D2">
	<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
	<td colspan=4 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#5582D2">
	<td colspan=4 valign="top" class="alt">&nbsp;<b>Vælg periode:</b></td>
</tr>
<tr><form action="joblog_afstem.asp?menu=kon&showall=1" method="post">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	<td width=50><b>Periode:</b></td><!--#include file="inc/weekselector_b.asp"-->
	<td align=right style="padding-right:45;"><input type="image" src="../ill/statpil.gif"></td>
<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
</tr>
<%
if len(request("FM_nuljob")) <> 0 then
nuljob = request("FM_nuljob")
njChk = "CHECKED"
else
nuljob = 0
njChk = ""
end if
%>
<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	<td colspan=4 valign="top">&nbsp;<input type="checkbox" name="FM_nuljob" id="FM_nuljob" value="1" <%=njChk%>>Vis "Nul" job. (Job uden registreringer)</td>
	<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
</tr>	
</form>
<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="5" alt="" border="0"></td>
		<td colspan=4 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="5" alt="" border="0"></td>
</tr>
</table>

<%
else%>
Periode: <b><%=formatdatetime(StrTdato, 1)%></b> - <b><%=formatdatetime(StrUdato, 1)%></b>
<%end if%>

<br><br>

<%
if request("print") <> "j" then
	if showall = 0 then%>
	<a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0"> Tilbage</a> 
	<img src="../ill/blank.gif" width="400" height="1" alt="" border="0"><a href="joblog_afstem.asp?print=j&fm_kunder=<%=selKunde%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" target="_blank" class=rmenu>Print venlig version <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<%else%>
	<img src="../ill/blank.gif" width="462" height="1" alt="" border="0"><a href="joblog_afstem.asp?print=j&fm_kunder=<%=selKunde%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" target="_blank" class=rmenu>Print venlig version <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<%end if%>
<%end if		


'*** Beskyt server mod belsatning mod langt datointerval ***
if datediff("d", sqlSTdato, sqlSLUTdato) > 95 OR datediff("d", sqlSTdato, sqlSLUTdato) < 0 then 'ca 3 md.
%><br><br>
<center>
<table cellpadding=0 cellspacing=0 border=0>			
		<tr>
			<td valign=top style="padding:5px; border:1px darkred solid;" bgcolor="#ffffe1">
			<img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0"><font class=error>Fejl!</font>
			<br><br><b>Der er 2 mulige årsager til at du modtager denne fejl:</b><br>
			<ul>
			<li>Det valgte datointerval er for stort. <br>
			Vælg et datointerval på mindre end 3 måneder. (95 dage)
			<br>
			<li>Det valgte dato-interval er negativt.
			</ul>
			</td>
		</tr>
</table>
</center>
<%else%>
		


		
		
		
					
		
			<%
			'**** Overblik Gand Grand total ****
			if showall = 1 then
			
			
			strSQL = "SELECT tknavn, tknr, kkundenr, kkundenavn, sum(t.timer) AS sumtimer, tjobnr, "_
			&" tfaktim, kunder.kid, tmnr FROM kunder "_
			&" LEFT JOIN timer t ON (t.tknr = kunder.kid AND "_
			&" tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim = 1)"_
			&" WHERE ketype <> 'e' GROUP BY kunder.kid ORDER BY kunder.kkundenavn" 
			
			'Response.write strSQl
			'Response.flush
			
			dim ikfbtimer
			dim fbtimer
			dim tknavn
			dim fakbeloebKunde
			dim faktimerKunde
			dim medfaktimer
			dim medfakbel
			dim tkid
			dim ialtregtimer
			dim tknr
			
			Redim ikfbtimer(0)
			Redim fbtimer(0)
			Redim tknavn(0)
			Redim faktimerKunde(0)
			Redim fakbeloebKunde(0)
			Redim medfaktimer(0)
			Redim medfakbel(0)
			Redim tkid(0)
			Redim ialtregtimer(0)
			Redim tknr(0)
			
			
			fakTimerKTot = 0
			fakBelKTot = 0
			medFakTimTot = 0
			medFakBelTot = 0
			
			oRec.open strSQL, oConn, 3 
			
			x = 0
			
			while not oRec.EOF 
			
			if lastKid <> oRec("kid") then
			lastKid = oRec("kid") 
			x = x + 1
			
			Redim preserve ikfbtimer(x)
			Redim preserve fbtimer(x)
			Redim preserve tknavn(x)
			Redim preserve faktimerKunde(x)
			Redim preserve fakbeloebKunde(x)
			Redim preserve medfaktimer(x)
			Redim preserve medfakbel(x)
			Redim preserve tkid(x)
			Redim preserve ialtregtimer(x)
			Redim preserve tknr(x)
			
			tkid(x) = oRec("kid")
			tknavn(x) = oRec("kkundenavn")
			tknr(x) = oRec("kkundenr")
			
			
			strSQL2 = "SELECT f.beloeb AS fakbeloeb, f.timer AS faktimer, "_
			&" sum(fm.fak) AS medfaktimer, sum(fm.beloeb) AS medfakbel, f.fakadr FROM fakturaer f, fak_med_spec fm WHERE "_
			&" (f.fakadr = "& oRec("kid") &" AND fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"') AND (fm.fakid = f.fid) GROUP BY fm.fakid"
			
			faktimerKunde(x) = 0
			fakbeloebKunde(x) = 0
			medfaktimer(x) = 0
			medfakbel(x) = 0
			
			'Response.write "<br><br>"& strSQL2
			'Response.flush
			
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF 
			faktimerKunde_this = faktimerKunde_this + oRec2("faktimer")
			fakbeloebKunde_this = fakbeloebKunde_this + oRec2("fakbeloeb") 
			medfaktimer_this = medfaktimer_this + oRec2("medfaktimer")
			medfakbel_this = medfakbel_this + oRec2("medfakbel")
			oRec2.movenext
			wend
			oRec2.close 
			
			
			
			faktimerKunde(x) = faktimerKunde_this
			fakbeloebKunde(x) = fakbeloebKunde_this
			medfaktimer(x) = medfaktimer_this
			medfakbel(x) = medfakbel_this
			
			faktimerKunde_this = 0
			fakbeloebKunde_this = 0
			medfaktimer_this = 0
			medfakbel_this = 0
			
				'*** Ikke fakbare timer på kunde ***
				strSQL3 = "SELECT sum(t.timer) AS sumIFBtimer, "_
				&" tfaktim FROM "_
				&" timer t WHERE t.tknr = "& oRec("kid") &" AND "_
				&" tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim = 2 GROUP BY t.tknr"
				
				'Response.write strSQL3
				'Response.flush
				
				ikfbtimer(x) = 0
				
				oRec3.open strSQL3, oConn, 3 
				if not oRec3.EOF then
					ikfbtimer(x) = oRec3("sumIFBtimer")
				end if 
				oRec3.close 
			 	
				ialtregtimer(x) = ialtregtimer(x) + ikfbtimer(x)
				regtimerTot = regtimerTot + ikfbtimer(x)
			
			end if 'LastKid <>
			
			
			
			
			fbtimer(x) = oRec("sumtimer")
			'** Grand Totaler ***
			regtimerTot = regtimerTot + fbtimer(x) 'ialtregtimer(x)
			fakbarregtimerTot = fakbarregtimerTot + fbtimer(x) 
			
			ialtregtimer(x) = ialtregtimer(x) + fbtimer(x)
			
			
			
			'Response.write tknavn(x) &": "& ialtregtimer(x) &" # "& fakbarregtimerTot & "<br>"
			
			'** Grand Totaler ***
			fakTimerKTot = fakTimerKTot + faktimerKunde(x)
			fakBelKTot = fakBelKTot + fakbeloebKunde(x)
			medFakTimTot = medFakTimTot + medfaktimer(x)
			medFakBelTot = medFakBelTot + medfakbel(x)
			
			'Response.write tknavn(x) &":"& medfaktimer(x) &" -- ialt:"& medFakTimTot &"<br>"
			
			oRec.movenext
			wend
			oRec.close 
			
			
			%>
			<table cellspacing=1 cellpadding=1 border=0 width=640>
			<tr>
				<td height=35>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td colspan=2 valign=bottom bgcolor="#ffffff" style="border:1px #999999 solid;">&nbsp;<img src="../ill/ac0009-24.gif" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Kontakter</b></td>
				<td colspan=2 valign=bottom bgcolor="#ffffff" style="border:1px #999999 solid;">&nbsp;<img src="../ill/ac0016-24.gif" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Medarbejdere</b></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td width=200 bgcolor="#ffffff" valign=bottom style="border:1px #999999 solid; padding-left:2px;" class=lille>Kontakt (kontakt id)</td>
				<td width=50 bgcolor="#e4e4e4" valign=bottom style="border:1px #999999 solid;" class=lille>Reg. timer ialt</td>
				<td width=50 bgcolor="#e4e4e4" valign=bottom style="border:1px #999999 solid;" class=lille>Heraf reg. fakturerbare timer</td>
				<td width=80 bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Tot. faktura timer</td>
				<td width=80 bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Tot. faktura beløb</td>
				<td width=80 bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Medarb. faktura timer</td>
				<td width=80 bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Medarb. faktura beløb</td>
				<td>&nbsp;</td>
			</tr>
			<%
			for x = 1 to x - 0
			
				if nuljob = 1 OR (ialtregtimer(x) <> 0 OR fbtimer(x) <> 0 OR faktimerKunde(x) <> 0 OR fakbeloebKunde(x) <> 0 OR medfaktimer(x) <> 0 OR medfakbel(x) <> 0) then %>
				<tr>
				<td bgcolor="#ffffff" style="border:1px #8caae6 solid; padding:2px;" class=lille>
				<%if request("print") <> "j" then%>
				<a href="joblog_afstem.asp?menu=kon&fm_kunder=<%=tkid(x)%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" target="_self" class=rmenu>
				<%=left(tknavn(x), 25)%>&nbsp;(<%=tknr(x)%>)</a>
				<%else%>
				<%=left(tknavn(x), 25)%>&nbsp;(<%=tknr(x)%>)
				<%end if%></td>
				<td align=right bgcolor="#e4e4e4" style="border:1px #999999 solid;" class=lille><%=formatnumber(ialtregtimer(x), 2)%> t.</td>
				<td align=right bgcolor="#e4e4e4" style="border:1px #999999 solid;" class=lille><%=formatnumber(fbtimer(x), 2)%> t.</td>
				<td align=right bgcolor="#ffffff" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(faktimerKunde(x), 2)%> t.</td>
				<td align=right bgcolor="#ffffff" style="border:1px limegreen solid;" class=lille><%=formatcurrency(fakbeloebKunde(x), 2)%></td>
				<td align=right bgcolor="#ffffff" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(medfaktimer(x), 2)%> t.</td>
				<td align=right bgcolor="#ffffff" style="border:1px limegreen solid;" class=lille><%=formatcurrency(medfakbel(x), 2)%></td>
				<td>
				<%
				
				if formatcurrency(fakbeloebKunde(x), 2) = formatcurrency(medfakbel(x), 2) then
				Response.write "<i><font color='limegreen'>&nbsp;V</font></i>"
				else
				Response.write "<i><font color='red'>&nbsp;-</font></i>"
				end if
				%>
				</td>
				</tr>
				<%
				end if
			next
			%>
			
			<tr>
			<td bgcolor="#ffffff" style="border:1px #8caae6 solid;" class=lille><b>Total:</b></td>
			<td align=right bgcolor="#ffffe1" style="border:1px #999999 solid;" class=lille><%=formatnumber(regtimerTot)%> t.</td>
			<td align=right bgcolor="#ffffe1" style="border:1px #999999 solid;" class=lille><%=formatnumber(fakbarregtimerTot)%> t.</td>
			<td align=right bgcolor="#ffffe1" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(fakTimerKTot)%> t.</td>
			<td align=right bgcolor="#ffffe1" style="border:1px limegreen solid;" class=lille><%=formatcurrency(fakBelKTot)%></td>
			<td align=right bgcolor="#ffffe1" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(medFakTimTot)%> t.</td>
			<td align=right bgcolor="#ffffe1" style="border:1px limegreen solid;" class=lille><%=formatcurrency(medFakBelTot)%></td>
			<td>&nbsp;</td>
			</tr>
			
			</table>
			<br><br>
			<%
			Call medtotaler(0)
			
			
			else
			
			'*** Viser detaljer på det valgte job%>
			
			<table cellpadding=0 cellspacing=1 border=0>			
			<tr>
			<td valign=top>
			<%
			totsumtimer = 0
			totsumfaktimer = 0
			totbel = 0
			totbelsamletfak = 0
			totfaktim = 0
			totantal = 0
			lastknr = 0
			lastknavn = ""
			x = 0
			
			totsumfaktimerprmed = 0
			totsumfakkrprmed = 0
			tottimerpaafak = 0
			totbelpaafak = 0
			
			tot_regtimer = 0
			tot_fakbaretimer = 0
			tot_fak_timer = 0
			tot_fak_belob = 0
			tot_med_fak_timer = 0
			tot_med_fak_belob = 0
			
			
			
			strSQL = "SELECT kkundenavn, tknr, kkundenr, jobnavn, tjobnr, job.id AS jid, "_
			&" jobnr, jobans1, jobans2, job.fastpris AS fastpris, jobTpris, budgettimer "_
			&" FROM kunder "_
			&" LEFT JOIN timer ON  (tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim = 1) "_
			&" LEFT JOIN job ON (jobknr = kunder.kid) "_ 
			&" WHERE "& selKundeKri &" AND ketype <> 'e' GROUP BY jid ORDER BY kkundenavn"
			
			oRec.open strSQL, oConn, 3 
			'Response.write strSQL
			
			while not oRec.EOF 
			
			if lastknr <> oRec("kkundenr") then
			%>
			<tr bgcolor="#5582d2">
				<td colspan=5 width=300 height=40 valign=bottom class=alt style="padding-left:5px; padding-top:5px; padding-bottom:5px; border-bottom:1px #5582d2 solid; border:1px #003399 solid;"><b><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)</b></td>
			</tr>
			<tr><td colspan=5 height=20>&nbsp;</td></tr>
			<tr><td colspan=2 class=lille valign=top>Navn</td><td class=lille valign=top>Registrerede timer</td><td class=lille valign=top>Registrerede fakturerbare timer og timepriser.</td><td class=lille valign=top>Faktura nr. timer og beløb.</td></tr>
			<%end if
			
					
					
					
					
						
						if oRec("fastpris") <> 1 then
						jtype = "Løbende timer"
						else
						jtype = "Fastpris"
						end if
						
						if oRec("jobans1") <> 0 then
						
						strSQL2 = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE mid = "& oRec("jobans1")
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txt = oRec2("mnavn") 
						jobans1nr = "&nbsp;("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						end if%>
						
						<%
						if oRec("jobans2") <> 0 then
						strSQL2 = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE mid = "& oRec("jobans2")
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") 
						jobans2nr = "&nbsp;("& oRec2("mnr") &")"
						end if
						oRec2.close
						
						end if
						
						
						
						if len(jobans1txt) <> 0 then 
						strJob = strJob & "Jobansvarlig 1:<br>"& left(jobans1txt, 18) &"&nbsp;"& jobans1nr &"<br>"
						end if
						if len(jobans2txt) <> 0 then 
						strJob = strJob & "Jobansvarlig 2:<br>"& left(jobans2txt, 18) &"&nbsp;"& jobans2nr
						end if
						
						jobans1txt = ""
						jobans2txt = ""
						jobans1nr = ""
						jobans2nr = ""
						
						
						if len(oRec("jid")) <> 0 then
						usejid = oRec("jid")
						else
						usejid = 0
						end if
						
								
								strSQL3 = "SELECT mnavn, sum(timer) AS sumtimer, tmnr, tmnavn, mid, mnr FROM medarbejdere RIGHT JOIN timer ON (tjobnr = '"& oRec("jobnr")&"' AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim <> 5 AND tmnr = mid) WHERE mid > 0 GROUP BY mid ORDER BY mid"
								
								oRec3.open strSQL3, oConn, 3 
								'Response.write strSQL3 & "<br>"
								while not oRec3.EOF 
								
									'*** Timer ***
									strTimer = "<tr bgcolor=#FFFFFF>"
									strTimer = strTimer &"<td valign=top colspan=2 style='padding-top:3; padding-left:5; border:1px #8caae6 solid;'>"&left(oRec3("mnavn"), 25)&"</td>"
									
									thissumtimer = oRec3("sumtimer")
									strTimer = strTimer & "<td align=right valign=top style='padding-top:3; padding-right:2px; padding-left:2px; border:1px #8caae6 solid;' class=lille>" & formatnumber(thissumtimer,2) &"</td>"
									totsumtimer = totsumtimer + thissumtimer
									
									'*****************************************************************
									'*** Fakbare timer ***
									'*****************************************************************
									strTimer = strTimer & "<td width=150 valign=top style='padding-top:3; padding-left:2px; border:1px #8caae6 solid;' class=lille>" 
									
									if oRec("fastpris") = 1 then '*Fastpris 
									strSQL4 = "SELECT sum(timer) AS sumfaktimer FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND tmnr = "& oRec3("mid") &" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim = 1" 
									else
									strSQL4 = "SELECT sum(timer) AS sumfaktimer, timepris FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND tmnr = "& oRec3("mid") &" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim = 1 GROUP BY timepris" 
									end if
									oRec4.open strSQL4, oConn, 3 
									
									while not oRec4.EOF 
									
									if cint(oRec("fastpris")) <> 1 then 
										if len(oRec4("sumfaktimer")) <> 0 then
										thissumfaktimer = oRec4("sumfaktimer")
										faktimertimepris = oRec4("timepris")
										else
										thissumfaktimer = 0
										faktimertimepris = 0
										end if
									else
										if len(oRec4("sumfaktimer")) <> 0 then
										thissumfaktimer = oRec4("sumfaktimer")
										faktimertimepris = (oRec("jobTpris") / oRec("budgettimer"))
										'Response.write "# "& oRec("jobTpris") & " - "& oRec("budgettimer")
										else
										thissumfaktimer = 0
										faktimertimepris = oRec("jobTpris")/1
										end if
									end if
									
									strTimer = strTimer & formatnumber(thissumfaktimer,2) &" á "
									strTimer = strTimer & formatcurrency(faktimertimepris, 2) & "<br>"
									
									totsumfaktimerprmed = totsumfaktimerprmed + thissumfaktimer
									totsumfakkrprmed = totsumfakkrprmed + (thissumfaktimer * faktimertimepris) 
									
									'*jobtotaler **
									totsumfaktimer = totsumfaktimer + thissumfaktimer
									totsumfakkr = totsumfakkr + (thissumfaktimer * faktimertimepris)
									oRec4.movenext
									wend 
									oRec4.close
									
									strTimer = strTimer & "Total: "& formatnumber(totsumfaktimerprmed, 2)&" timer - "& formatcurrency(totsumfakkrprmed, 2) &"</td>"
									
									
									totsumfaktimerprmed = 0
									totsumfakkrprmed = 0
									'******************************************************************
									
									
									strTimer = strTimer & "<td valign=top style='border:1px #8caae6 solid;'>"
									
									strFak = "<table cellspacing=0 cellpadding=0 border=0>"
									
									if len(oRec3("mid")) <> 0 then
									usemid = oRec3("mid")
									else
									usemid = 0
									end if
									
									
									 
									
									
									'**********************************************************************
									'*** Fakturaer, fakturerede timer og beløb ***
									'**********************************************************************
									strSQL4 = "SELECT jobid, faknr, fid FROM fakturaer"_
									&" WHERE jobid = "& usejid &" AND faktype = 0 AND fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' ORDER BY fid " 
									oRec4.open strSQL4, oConn, 3 
									
									while not oRec4.EOF 
										
										strFak = strFak &"<tr><td width=50 valign=top class=lille style='padding-left:2px;'><font color='#2c962d'>"& oRec4("faknr") &"</font></td>"
										
										strSQL2 = "SELECT sum(fak) AS medsumfaktimer, enhedspris FROM fak_med_spec"_
										&" WHERE fakid = "& oRec4("fid") &" AND mid = " & usemid &" AND fak <> 0 GROUP BY enhedspris"
										oRec2.open strSQL2, oConn, 3 
										
										strFak = strFak & "<td class=lille valign=top width=130><font color='#5582d2'>"
										while not oRec2.EOF 
											
											if len(oRec2("medsumfaktimer")) <> 0 then
											faktureredetimer = oRec2("medsumfaktimer")
											else
											faktureredetimer = 0
											end if
											
											fakenhpris = oRec2("enhedspris")
											
											strFak = strFak & formatnumber(faktureredetimer, 2) &" á "& formatcurrency(fakenhpris, 2) &"<br>"
											faktureredetimerialt = faktureredetimerialt + faktureredetimer
										oRec2.movenext
										wend
										oRec2.close
										
										strFak = strFak & "</td>"
									
									
									
										strSQL2 = "SELECT sum(beloeb) AS bel FROM fak_med_spec "_
										&" WHERE fakid = "& oRec4("fid") &" AND mid = " & usemid &" AND fak <> 0"
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then
											bel = oRec2("bel")
										end if
										oRec2.close
										
											if len(bel) <> 0 then
											bel = bel 
											else
											bel = 0
											end if
									
									
									strFak = strFak & "<td width=40 class=lille align=right valign=top style='padding-right:2;'>"& formatnumber(faktureredetimerialt) &"</td>"
									strFak = strFak & "<td width=60 class=lille align=right valign=top style='padding-right:2;'>"& formatcurrency(bel) & "</td></tr>"
									
									
									'**** faktotaler på medareb ***
									faktimertotalpamedpafak = faktimertotalpamedpafak + faktureredetimerialt
									fakkrprmedprfak = fakkrprmedprfak + bel
									
									'*** faktotaler på job ***
									totbelsamletfak = totbelsamletfak + bel
									totfaktim = totfaktim + faktureredetimerialt
									
									faktureredetimerialt = 0
									
									oRec4.movenext
									wend
									oRec4.close 
									
									
									'*** faktureret totaler på medarbejder ***
									strFak = strFak & "<tr>"
									strFak = strFak & "<td colspan=2 width=180 align=right class=lille style='padding-right:2px;'>Total:</td>"
									strFak = strFak & "<td align=right width=40 class=lille style='padding-right:2px;'>"&formatnumber(faktimertotalpamedpafak)&"</td>"
									strFak = strFak & "<td align=right width=60 class=lille style='padding-right:2px;'>"&formatcurrency(fakkrprmedprfak)&"</td>"
									strFak = strFak & "</tr>"
									strFak = strFak & "</table>"
									
									strTimer2 = "</td>"
									strTimer2 = strTimer2 & "</tr>"
								
								faktimertotalpamedpafak = 0
								fakkrprmedprfak = 0
								'*************************************************
								
								'*** Udskriver Timer og Faktimer på medarbjeder hvis de findes ***
									if thissumtimer > 0 OR faktureredetimer > 0 then
										
										'if oRec("jid") <> lastjid then 	
										'Response.write strJob & oRec("jid") & "<br>"
										'end if
										'strJob = ""
										
									Response.write strTimer
									Response.write strFak
									Response.write strTimer2
									end if
									
									strTimer = ""
									strFak = ""
									strTimer2 = ""
									
									faktureredetimer = 0
									thissumtimer = 0
								
								oRec3.movenext
								wend
								oRec3.close 
								
								
								
								'*****************************************************
								'*** Jobtotaler ***
								'*****************************************************
								
								if len(totsumtimer) <> 0 then
								totsumtimer = totsumtimer
								else
								totsumtimer = 0
								end if
								
								if len(totsumfaktimer) <> 0 then
								totsumfaktimer = totsumfaktimer
								else
								totsumfaktimer = 0
								end if
								
								if len(totsumfakkr) <> 0 then
								totsumfakkr = totsumfakkr
								else
								totsumfakkr = 0
								end if
								
								
									'*** total beløb på fak uafhængig af medarb ***
									strSQL4 = "SELECT sum(beloeb) AS totbeloebpaafak FROM fakturaer"_
									&" WHERE jobid = "& usejid &" AND faktype = 0 AND fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' ORDER BY fid " 
									
									'Response.write strSQL4
									'Response.flush
									
									oRec4.open strSQL4, oConn, 3 
									if not oRec4.EOF then
										totbelpaafak = totbelpaafak + oRec4("totbeloebpaafak")
									end if
									oRec4.close
									
									'*** total timer på fak uafhængig af medarb ***
									strSQL4 = "SELECT sum(timer) AS tottimerpaafak FROM fakturaer"_
									&" WHERE jobid = "& usejid &" AND faktype = 0 AND fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' ORDER BY fid " 
									oRec4.open strSQL4, oConn, 3 
									if not oRec4.EOF then
										tottimerpaafak = tottimerpaafak + oRec4("tottimerpaafak")
									end if
									oRec4.close
									
									if len(tottimerpaafak) <> 0 then
									tottimerpaafak = tottimerpaafak
									else
									tottimerpaafak = 0
									end if
									
									if len(totbelpaafak) <> 0 then
									totbelpaafak = totbelpaafak
									else
									totbelpaafak = 0
									end if
								 %>
								
								<%if tottimerpaafak <> 0 OR totsumtimer <> 0 then%>
								<tr bgcolor="#e4e4e4" height=25>
									<td colspan=5 style="padding-left:5px; border:1px #8caae6 solid;"><b><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</b>
								</tr>
								<tr bgcolor="#ffffe1">	
									<td bgcolor="#e4e4e4" colspan=2 style="padding-left:5px; border:1px #8caae6 solid;" valign=bottom width=150 height=55>
									<font class=megetlillesort>
									Jobtype: <%=jtype%><br>
									<%=strJob%>
									</td>
									<td align=right style="padding-right:2px; padding-left:2px; border:1px #8caae6 solid;" valign=bottom class=lille><%=formatnumber(totsumtimer, 2)%></td>
									<td style="padding-left:2px; border:1px #8caae6 solid;" valign=bottom class=lille>Ialt: <%=formatnumber(totsumfaktimer, 2)%> timer - <%=formatcurrency(totsumfakkr)%></td>
									<td valign=bottom style="border:1px #8caae6 solid;"> 
									<table cellspacing=0 cellpadding=0 border=0>
									<tr>
										<td width=180 valign=bottom style="padding-left:2px;" class=lille><font color="#2c962d">(<%=formatnumber(tottimerpaafak)%> timer <%=formatcurrency(totbelpaafak)%>)</font></td>
										<td width=40 valign=bottom align=right style="padding-right:2;" class=lille><%=formatnumber(totfaktim)%></td>
										<td width=60 valign=bottom align=right style="padding-right:2;" class=lille><%=formatcurrency(totbelsamletfak)%></td>
									</tr>
									</table>
									</td>
									</tr>
									<tr><td colspan=5 height=25>&nbsp;</td></tr>
								<%end if
									
					strJob = ""
					totsumtimer = 0 'ok
					
					tottimerpaafak = 0'ok
					totbelpaafak = 0 'ok
					totsumfaktimer = 0 'ok
					totsumfakkr = 0 'ok
					
					totbel = 0
					totfaktim = 0
					
					totbelsamletfak = 0
					
					
					
					
					
			
			
			lastknr = oRec("kkundenr")
			lastknavn = oRec("kkundenavn")
			x = x + 1
			Response.flush
			oRec.movenext
			wend
			oRec.close 
			
			%>
			</table>
			</table>
			<%
			
			if x <> 0 then
				call medtotaler(selKunde)
			end if
			%>
		
		
		<%
		end if '*** Vis alle/Vis detaljer
		
		end if '** Over 95 dage valgt!%>
		</div>	
<%
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
