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
		
    menu = "erp"
	print = request("print")
		
		
    if print <> "j" then%>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <!--#include file="../inc/regular/topmenu_inc.asp"-->
    <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	    <%call erpmainmenu(2)%>
	    </div>
	    <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	    <%
	    call erptopmenu2()
	    %>
	    </div>

    <%else%>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
    <%end if
		
	'*** Finder akttyper der er fakturerbare, ikke fakbare / tælles med i timereg ***	
	call akttyper2009(2)	
		
	public totfaktimer, totfakbelob, totmedtimer, totmedbel 
	function medarbejderTD(lastMid, thismedtimer, thismedbelob, kundeKri2, strMedNavn)
				
				strMedTimerogFak = "<td width=80 bgcolor=#ffffff align=right class=lille style='border:1px #999999 solid;'>"&formatnumber(thismedtimer)&" t.</td>"
				strMedTimerogFak = strMedTimerogFak &"<td class=lille bgcolor=#ffffff width=80 align=right style='border:1px #999999 solid;'>"&formatnumber(thismedbelob)&" DKK</td>"
				
				strSQL2 = "SELECT SUM(fm.fak) AS medfaktimer, SUM(fm.beloeb) AS medfakbel, "_
				&" f.faktype, f.kurs AS fakkurs, f.valuta, fm.valuta, fm.kurs AS fmskurs "_
				&" FROM fakturaer f "_
				&" LEFT JOIN fak_med_spec fm ON (fm.fakid = f.fid AND fm.mid = "& lastMid &") WHERE "_
				&" (fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' "& kundeKri2 &") AND fm.enhedsang = 0 "_
				&" GROUP BY f.faktype, fm.mid, fm.valuta, fm.kurs"
				
				'Response.write strSQL2 &"<br><br>"
				'Response.flush
				
				medFakTimerTot = 0
				medFakBelTot = 0
				nyMedarbFakbel = 0
				
				oRec2.open strSQL2, oConn, 3 
				while not oRec2.EOF 
				
				
				'** Beløb i FMS tabel er korrigeret for KURS ***
				nyMedarbFakbel = oRec2("medfakbel")
		        
				call beregnValuta(nyMedarbFakbel,oRec2("fakkurs"),100)
			    nyMedarbFakbel = valBelobBeregnet
				
				if oRec2("faktype") <> 1 then
				medFakTimerTot = medFakTimerTot + oRec2("medfaktimer")
				medFakBelTot = medFakBelTot + nyMedarbFakbel
				else
				medFakTimerTot = medFakTimerTot - (oRec2("medfaktimer"))
				medFakBelTot = medFakBelTot - (nyMedarbFakbel)
				end if
				
				oRec2.movenext
				wend
				oRec2.close 
			    
			    
			    if len(medFakTimerTot) <> 0 then
			    medFakTimerTot = medFakTimerTot
			    else
			    medFakTimerTot = 0
			    end if
			    
			    if len(medFakBelTot) <> 0 then
			    medFakBelTot = medFakBelTot
			    else
			    medFakBelTot = 0
			    end if
				
				strMedTimerogFak = strMedTimerogFak &"<td bgcolor=#ffffff class=lille align=right style='border:1px #999999 solid;'>"&formatnumber(medFakTimerTot)&" t.</td>"
				strMedTimerogFak = strMedTimerogFak & "<td class=lille bgcolor=#ffffff align=right style='border:1px #999999 solid;'>"&formatnumber(medFakBelTot)&" DKK</td></tr>"
				
				
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
	end function
	
	function medtotaler(sallkid)
	%><br><br>
	<img src="../ill/ac0016-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Medarbejder totaler:</b><br>
			<table cellpadding=1 cellspacing=1 border=0 width="800">
			<tr>
				<td bgcolor="#ffffff" valign=bottom class=lille style="padding-left:2px; border:1px #999999 solid;">Navn</td>
				<td bgcolor="#ffffff" colspan=2 class=lille style="border:1px #999999 solid;">Registreret <br>(kun fakturerbare)</td>
				<td bgcolor="#ffffff" colspan=2 class=lille valign=bottom style="padding-left:2px; border:1px #999999 solid;">Timer faktureret<br />
				(stk, enhed og km bliver ikke medregnet)</td>
			</tr>
			<%
			
			'if sallkid = 0 then
			'jtype = "LEFT"
			'kundeKri1 = ""
			'kundeKri2 = ""
			'else
			'jtype = "LEFT"
			if kundeid <> 0 then
			kundeKri1 = "  t.tknr = " & kundeid & " AND "
			kundeKri2 = " AND fakadr = " & kundeid 
			else
			kundeKri1 = "  t.tknr <> " & kundeid & " AND "
			kundeKri2 = " AND fakadr <> " & kundeid 
			end if
			'end if
			
			
	        
			
			
			strSQL = "SELECT medarbejdere.mid AS medid, mnr, mnavn, sum(t.timer) AS sumtimer, "_
			&" timepris, j.fastpris, j.budgettimer, j.ikkebudgettimer, jobtpris, t.valuta, t.kurs, "_
			&" usejoborakt_tp, a.aktbudget FROM medarbejdere "_
			&" LEFT JOIN timer t ON ("& kundeKri1 &" tmnr = medarbejdere.mid AND ("& aty_sql_realHoursFakbar &") "_
			&" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"')"_
			&" LEFT JOIN job j ON (jobnr = tjobnr) "_
			&" LEFT JOIN aktiviteter AS a ON (a.id = t.taktivitetid) "_
			&" GROUP BY medarbejdere.mid, taktivitetid, jobnr, kurs, valuta ORDER BY mnavn"
			
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
			
			strMedNavn = "<tr><td bgcolor=#ffffff class=lille style='padding-left:2px; border:1px #999999 solid;'>"&oRec("mnavn")&" ("&oRec("mnr")&")</td>"
			end if%>
			
			<%
			thismedtimer = thismedtimer + oRec("sumtimer")
			
			if len(thismedtimer) <> 0 then
			thismedtimer = thismedtimer
			else
			thismedtimer = 0
			end if
			
			'Response.write oRec("medid") & ": "& thismedtimer & "<br>"
			
			
			if oRec("budgettimer") <> 0 then
			budgettimer = oRec("budgettimer")
			else
			budgettimer = 1
			end if
			
			nytBeregnetBelob = 0
			
			if oRec("fastpris") <> 1 then
			nytBeregnetBelob = (oRec("sumtimer") * oRec("timepris"))
			else
			    '*** Job eller akt er grundlag
			    if oRec("usejoborakt_tp") <> 0 then
    			nytBeregnetBelob = (oRec("sumtimer") * oRec("aktbudget"))
			    else
			    nytBeregnetBelob = (oRec("sumtimer") * (oRec("jobtpris")/budgettimer))
			    end if
			end if
			
			
			call beregnValuta(nytBeregnetBelob,oRec("kurs"),100)
		    nytBeregnetBelob = valBelobBeregnet
		    
		    thismedbelob = thismedbelob + nytBeregnetBelob
		    
		    
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
				<td bgcolor="lightpink" class=lille style="border:1px #999999 solid;"><b>Total:</b></td>
				<td bgcolor="lightpink" align=right class=lille style="border:1px #999999 solid;"><%=formatnumber(totmedtimer)%> t.</td>
				<td bgcolor="lightpink" align=right class=lille style="border:1px #999999 solid;"><%=formatnumber(totmedbel, 2)%> DKK</td>
				<td bgcolor="lightpink" align=right class=lille style="border:1px #999999 solid;"><%=formatnumber(totfaktimer)%> t.</td>
				<td bgcolor="lightpink" align=right class=lille style="border:1px #999999 solid;"><%=formatnumber(totfakbelob, 2)%> DKK</td>
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
				    select case strAar
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    strDag = 29
				    case else
				    strDag = 28
				    end select
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
				    
		    select case strAar_slut
		    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
		    strDag_slut = 29
		    case else
		    strDag_slut = 28
		    end select
				    
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
		if len(Request.Cookies("erp_datoer")("st_dag")) <> 0 then
	
		strMrd = Request.Cookies("erp_datoer")("st_md")
		strDag = Request.Cookies("erp_datoer")("st_dag")
		strAar = Request.Cookies("erp_datoer")("st_aar") 
		strDag_slut = Request.Cookies("erp_datoer")("sl_dag")
		strMrd_slut = Request.Cookies("erp_datoer")("sl_md")
		strAar_slut = Request.Cookies("erp_datoer")("sl_aar")
		
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
	Response.Cookies("erp_datoer")("st_dag") = strDag
	Response.Cookies("erp_datoer")("st_md") = strMrd
	Response.Cookies("erp_datoer")("st_aar") = strAar
	Response.Cookies("erp_datoer")("sl_dag") = strDag_slut
	Response.Cookies("erp_datoer")("sl_md") = strMrd_slut
	Response.Cookies("erp_datoer")("sl_aar") = strAar_slut
	Response.Cookies("erp_datoer").Expires = date + 10		
	
	
	'**** SQL datoer ***
	sqlSTdato = "20" & strAar &"/"& strMrd &"/"& strDag 
	sqlSLUTdato = "20" & strAar_slut &"/"& strMrd_slut &"/"& strDag_slut 
	
	
	if request("print") <> "j" then
	dleft = "20"
	dtop = "132"
	
	else
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="800">
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
		<%oimg = "view_1_1.gif"
	oleft = 0
	otop = 0
	owdt = 700
	oskrift = "Afstemning job og medarbejdere (timer realiseret / faktureret)"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
		
		
	

 call filterheader(0,0,800,pTxt)
 
	'*** Valgt kunde ***
	if len(request("FM_kunder")) <> 0 AND request("FM_kunder") <> 0 then
	selKunde = request("FM_kunder")
	selKundeKri = " kunder.kid = " & selKunde &""
	    
	    if len(trim(request("showall"))) <> 0 AND request("showall") <> 0 then
	    showall = 1
	    else
	    showall = 0
	    end if
	    
	'showall = 0
	else
	selKunde = 0
	selKundeKri = " kunder.kid <> 0"
	showall = 1
	end if
	
	kundeid = selKunde
	
	

if request("print") <> "j" AND showall = 1 then%>

<table cellspacing="0" cellpadding="0" border="0" width="100%">


<tr><form action="erp_job_afstem.asp?menu=erp&showall=1" method="post">
    <td colspan=3><b>Kontakter (kunder):</b><br />
    
    <%
    
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid, COUNT(f.fid) AS antalfak FROM kunder "_
	    &" LEFT JOIN fakturaer f ON (f.fakadr = kid AND shadowcopy = 0) "_
		&" WHERE ketype <> 'e'"_
		&" AND f.fid <> 0 GROUP BY kid ORDER BY Kkundenavn"
		'Response.write strSQL & "<br>"
		'Response.flush		
	
	
	if print <> "j" AND media <> "export" then
		%>
		<select name="FM_kunder" id="FM_kunder" style="font-size : 11px; width:406px;" onChange="submit();">
		<option value="0">Alle - eller vælg fra liste...</option>
		<%
	end if
				
				oRec.open strSQL, oConn, 3
				k = 0
				while not oRec.EOF
				
				
				if print <> "j" AND media <> "export" then
				
				if cint(kundeid) = cint(oRec("kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>) - antal. fak. ialt: <%=oRec("antalfak") %></option>
				<%
				
				end if
				
				k = k + 1
				oRec.movenext
				wend
				oRec.close
				
				
				if print <> "j" AND media <> "export" then
				%>
				</select>
		        <%end if
    %>
    <br />&nbsp;
    </td>
</tr>
<tr>
	<td style="white-space:nowrap; width:100px;"><b>Vælg periode:</b></td><!--#include file="inc/weekselector_b.asp"-->
	<td align=right style="padding-right:45;">
        <input id="Submit1" type="submit" value=" Kør >> " /></td>
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
	<td colspan=3>&nbsp;<input type="checkbox" name="FM_nuljob" id="FM_nuljob" value="1" <%=njChk%>>Vis "Nul" job. (Job uden registreringer)</td>
	</tr>	
</form>

</table>

<%
else%>
Periode: <b><%=formatdatetime(StrTdato, 1)%></b> - <b><%=formatdatetime(StrUdato, 1)%></b>
<%end if%>
<!-- filterDiv -->
	</td></tr>
	</table>
	</div>
	
<br><br>
<%
 if showall = 0 then%>
	        <a href="Javascript:history.back()" class=vmenu><img src="../ill/arrow_left_blue.png" alt="Tilbage" border="0"> </a> 
	       <%end if%>

<%		
'*** Beskyt server mod belsatning mod langt datointerval ***
if datediff("d", sqlSTdato, sqlSLUTdato) > 365 OR datediff("d", sqlSTdato, sqlSLUTdato) < 0 then 'ca 3 md.
%><br><br>
<center>
<table cellpadding=0 cellspacing=0 border=0>			
		<tr>
			<td valign=top style="padding:20px; border:2px darkred dashed;" bgcolor="#ffff99">
			<img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0"><font class=error>Fejl!</font>
			<br><br><b>Der er 2 mulige årsager til at du modtager denne fejl:</b><br>
			<ul>
			<li>Det valgte datointerval er for stort. <br>
			Vælg et datointerval på mindre end 12 måneder. (365 dage)
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
			
			
			strSQL = "SELECT kkundenr, kkundenavn, "_
			&" kunder.kid FROM kunder "_
			&" WHERE "& selKundeKri &" AND ketype <> 'e' GROUP BY kunder.kid ORDER BY kunder.kkundenavn" 
			
			'Response.write strSQl
			'Response.flush
			
			'dim ikfbtimer
			dim realtimer
			dim fbtimer
			dim tknavn
			dim fakbeloebKunde
			dim faktimerKunde
			dim medfaktimer
			dim medfakbel
			dim tkid
			'dim ialtregtimer
			dim tknr
			
			'Redim ikfbtimer(0)
			redim realtimer(350)
			Redim fbtimer(350)
			Redim tknavn(350)
			Redim faktimerKunde(350)
			Redim fakbeloebKunde(350)
			Redim medfaktimer(350)
			Redim medfakbel(350)
			Redim tkid(350)
			'Redim realtimer(350)
			Redim tknr(350)
			
			
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
			
			
			
			tkid(x) = oRec("kid")
			tknavn(x) = oRec("kkundenavn")
			tknr(x) = oRec("kkundenr")
			
			
			realtimer(x) = 0
			
			'**** Alle timer *****
			strSQL2 = "SELECT sum(t2.timer) AS realtimer FROM "_
			&" timer t2 WHERE t2.tknr = "& tkid(x) &" AND "_
			&" t2.tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND ("& replace(aty_sql_realHours, "tfak", "t2.tfak") &")"_
			&" GROUP BY tknr"
			
			oRec2.open strSQL2, oConn, 3
			if not oRec2.EOF then
			
			realtimer(x) = oRec2("realtimer")
			
			end if
			oRec2.close
			
			
			fbtimer(x) = 0
			
			'**** Fakbare timer *****
			strSQL2 = "SELECT sum(timer) AS fbtimer FROM "_
			&" timer t2 WHERE t2.tknr = "& tkid(x) &" AND "_
			&" t2.tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND ("& replace(aty_sql_realHoursFakbar, "t.", "t2.") &")"_
			&" GROUP BY tknr"
			
			oRec2.open strSQL2, oConn, 3
			if not oRec2.EOF then
			
			fbtimer(x) = oRec2("fbtimer")
			
			end if
			oRec2.close
			
			
			'*** Faktura medab. linier ***'
			'*** Kun dem med enhed = 0 (timer) ***'
			strSQL2 = "SELECT f.faktype, "_
			&" fm.fak AS medfaktimer, fm.beloeb AS medfakbel, "_
			&" f.kurs AS fakkurs, fm.kurs AS fmskurs FROM "_
			&" fakturaer f LEFT JOIN fak_med_spec fm ON (fm.fakid = f.fid) "_
			&" WHERE "_
			&" (f.fakadr = "& oRec("kid") &" AND fakdato BETWEEN '"& sqlSTdato &"' "_
			&" AND '" & sqlSLUTdato &"') AND fm.enhedsang = 0 "_
			&" ORDER BY fm.fakid, fm.kurs, fm.valuta, fm.mid"
			'&" GROUP BY f.faktype, fm.fakid, fm.kurs, fm.valuta, fm.mid"
			
			medfaktimer(x) = 0
			medfakbel(x) = 0
			nyMedarbFakbel = 0
			
			'Response.write "<br><br>"& strSQL2
			'Response.flush
			
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF 
			
			call beregnValuta(oRec2("medfakbel"),oRec2("fakkurs"),100)
			nyMedarbFakbel = valBelobBeregnet
			
			'Response.Write oRec2("medfakbel") & " - " & oRec2("fakkurs") & "<br>"
			
			if oRec2("faktype") <> 1 then
			medfaktimer_this = medfaktimer_this + oRec2("medfaktimer")
			medfakbel_this = medfakbel_this + nyMedarbFakbel
			else
			medfaktimer_this = medfaktimer_this -(oRec2("medfaktimer"))
			medfakbel_this = medfakbel_this -(nyMedarbFakbel)
			end if
			
			oRec2.movenext
			wend
			oRec2.close 
			
			
			
			'*** Fak tot **'
			strSQL2 = "SELECT f.beloeb AS fakbeloeb, f.faktype, "_
			&" f.fakadr, sum(fd.antal) AS faktimer, "_
			&" f.kurs AS fakkurs FROM "_
			&" fakturaer f "_
			&" LEFT JOIN faktura_det fd ON (fd.fakid = f.fid) WHERE "_
			&" (f.fakadr = "& oRec("kid") &" AND "_
			&" fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"') AND fd.enhedsang = 0  "_
			&" GROUP BY f.faktype, f.fid "
			
			'Response.Write strSQL2
			'Response.flush
			
			faktimerKunde(x) = 0
			fakbeloebKunde(x) = 0
			nyFakbel = 0
			
		    'Response.write "<br><br>"& strSQL2
			'Response.flush
			
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF 
			
			call beregnValuta(oRec2("fakbeloeb"),oRec2("fakkurs"),100)
			nyFakbel = valBelobBeregnet
			
			if oRec2("faktype") <> 1 then
			faktimerKunde_this = faktimerKunde_this + oRec2("faktimer")
			fakbeloebKunde_this = fakbeloebKunde_this + nyFakbel 
			else
			faktimerKunde_this = faktimerKunde_this -(oRec2("faktimer"))
			fakbeloebKunde_this = fakbeloebKunde_this -(nyFakbel) 
			end if
			
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
			   
			    
			   
			end if 'LastKid <>
			
			
			'realtimer(x) = 0
			
			
			
			'** Grand Totaler ***
			regtimerTot = regtimerTot + realtimer(x) 'realtimer(x)
			fakbarregtimerTot = fakbarregtimerTot + fbtimer(x) 
		    'realtimer(x) = realtimer(x) + fbtimer(x)
			'Response.write tknavn(x) &": "& realtimer(x) &" # "& fakbarregtimerTot & "<br>"
			
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
           
            
			
			
			<table cellspacing=1 cellpadding=1 border=0 width="1004">
			<tr>
				<td height=35>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td colspan=2 valign=bottom bgcolor="#ffffff" style="border:1px #999999 solid;">&nbsp;<img src="../ill/ikon_kunder_24.png" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Kontakter</b></td>
				<td colspan=2 valign=bottom bgcolor="#ffffff" style="border:1px #999999 solid;">&nbsp;<img src="../ill/ac0016-24.gif" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Medarbejdere</b></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td bgcolor="#ffffff" valign=bottom style="border:1px #999999 solid; padding-left:2px;" class=lille>Kontakt (kontakt id)</td>
				<td bgcolor="#e4e4e4" valign=bottom style="border:1px #999999 solid;" class=lille>Reg. timer ialt</td>
				<td bgcolor="#e4e4e4" valign=bottom style="border:1px #999999 solid;" class=lille>Heraf reg. fakturerbare timer</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Tot. fak. timer <br />incl. "herreløse" (sum-akt.)
				<br />(stk, enhed og km bliver ikke medregnet)</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Tot. faktura beløb</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Medarb. fak. timer
				<br />(stk, enhed og km bliver ikke medregnet)</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border:1px #999999 solid;" class=lille>Medarb. faktura beløb</td>
				<td>&nbsp;</td>
			</tr>
			<%
			for x = 1 to x - 0
			
				if nuljob = 1 OR (realtimer(x) <> 0 OR fbtimer(x) <> 0 OR faktimerKunde(x) <> 0 OR fakbeloebKunde(x) <> 0 OR medfaktimer(x) <> 0 OR medfakbel(x) <> 0) then %>
				<tr>
				<td bgcolor="#ffffff" style="border:1px #8caae6 solid; padding:2px;" class=lille>
				<%if request("print") <> "j" then%>
				<a href="erp_job_afstem.asp?menu=erp&fm_kunder=<%=tkid(x)%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" target="_self" class=rmenu>
				<%=left(tknavn(x), 25)%>&nbsp;(<%=tknr(x)%>)</a>
				<%else%>
				<%=left(tknavn(x), 25)%>&nbsp;(<%=tknr(x)%>)
				<%end if%></td>
				<td align=right bgcolor="#e4e4e4" style="border:1px #999999 solid;" class=lille><%=formatnumber(realtimer(x), 2)%> t.</td>
				<td align=right bgcolor="#e4e4e4" style="border:1px #999999 solid;" class=lille><%=formatnumber(fbtimer(x), 2)%> t.</td>
				<td align=right bgcolor="#ffffff" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(faktimerKunde(x), 2)%> t.</td>
				<td align=right bgcolor="#ffffff" style="border:1px limegreen solid;" class=lille><%=formatnumber(fakbeloebKunde(x), 2)%> DKK</td>
				<td align=right bgcolor="#ffffff" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(medfaktimer(x), 2)%> t.</td>
				<td align=right bgcolor="#ffffff" style="border:1px limegreen solid;" class=lille><%=formatnumber(medfakbel(x), 2)%> DKK</td>
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
			<td bgcolor="lightpink" style="border:1px #8caae6 solid;" class=lille><b>Total:</b></td>
			<td align=right bgcolor="lightpink" style="border:1px #999999 solid;" class=lille><%=formatnumber(regtimerTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border:1px #999999 solid;" class=lille><%=formatnumber(fakbarregtimerTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(fakTimerKTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border:1px limegreen solid;" class=lille><%=formatnumber(fakBelKTot, 2)%> DKK</td>
			<td align=right bgcolor="lightpink" style="border:1px #8caae6 solid;" class=lille><%=formatnumber(medFakTimTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border:1px limegreen solid;" class=lille><%=formatnumber(medFakBelTot, 2)%> DKK</td>
			<td>&nbsp;</td>
			</tr>
			
			</table>
			<br><br>
			<%
			call medtotaler(0)
			
			
			else
			
			'****************************************'
			'*** Viser detaljer på det valgte job ***'
			'****************************************'
			
			
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
			
			
			'faktimerPrMed = 0
			fakmatBeloeb = 0
			'fakmedarbBeloeb = 0
			fakaktBeloeb = 0
			fakaktKorsBeloeb = 0
			
			
			strSQL = "SELECT kkundenavn, kkundenr, adresse, postnr, city, land, jobnavn, job.id AS jid, "_
			&" jobnr, jobans1, jobans2, job.fastpris AS fastpris, jobTpris, budgettimer, usejoborakt_tp "_
			&" FROM kunder "_
			&" LEFT JOIN job ON (jobknr = kunder.kid) "_ 
			&" WHERE "& selKundeKri &" AND ketype <> 'e' GROUP BY jid ORDER BY kkundenavn"
			
			'&" LEFT JOIN timer ON (tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND tfaktim = 1) "_
			'job.id AS jid, jobnr,
			
			oRec.open strSQL, oConn, 3 
			'Response.write "<br>"& strSQL
			
			%>
			
			<table cellpadding=2 cellspacing=2 border=0 width="800">			
			<tr>
			<td valign=top>
			<%
			
			
			while not oRec.EOF 
			
			if lastknr <> oRec("kkundenr") then
			%>
			<tr>
				<td colspan=5 height=40 valign=bottom style="padding:5px 5px 5px 5px;">
                    <img src="../ill/ikon_kunder_24.png" />&nbsp;
                    <b><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)</b><br />
                    <%=oRec("adresse") %><br />
                    <%=oRec("postnr") %>, <%=oRec("city") %><br />
                    <%=oRec("land") %></td>
			</tr>
			<tr><td colspan=5 height=20>&nbsp;</td></tr>
			<tr bgcolor="#5582d2"><td colspan=2 valign=top class=alt>Navn</td>
			<td class=alt valign=top>Registrerede timer</td>
			<td class=alt valign=top>Registrerede fakturerbare timer og timepriser.</td>
			<td class=alt valign=top>Faktura nr., fakturerede timer og beløb pr. medarb.<br />
			(kun timer, ikke stk. enheder og km.)</td></tr>
			<%end if
			
			
			
			strthisJob = "<tr bgcolor=""#e4e4e4"" height=""25"">"_
			&"<td colspan=""5"" style=""padding:5px 5px 5px 5px; border:1px #8caae6 solid;"">"_
            &"<img src=""../ill/ikon_job_24.png"" />&nbsp;"_
            &"<b>"&oRec("jobnavn")&" ("&oRec("jobnr")&")</b></td></tr>"
			
			
					
					
					
					
						
						if oRec("fastpris") <> 1 then
						jtype = "Lbn. timer"
						else
						    if oRec("usejoborakt_tp") <> 0 then
						    jtype = "Fastpris (akt.)"
						    else
						    jtype = "Fastpris (job)"
						    end if
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
						strJob = strJob & "Jobansvarlig:<br>"& left(jobans1txt, 18) &"&nbsp;"& jobans1nr &"<br>"
						end if
						if len(jobans2txt) <> 0 then 
						strJob = strJob & "Jobejer:<br>"& left(jobans2txt, 18) &"&nbsp;"& jobans2nr
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
						
						
						
								
								'**** Alle timer (alle akt. typer) på medarb ***'
								strSQL3 = "SELECT mnavn, sum(timer) AS sumtimer, tmnr, tmnavn, mid, mnr FROM medarbejdere LEFT JOIN timer ON (tjobnr = "&oRec("jobnr")&" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND "_
								&" ("& aty_sql_realhours &") AND tmnr = mid) WHERE mid > 0 GROUP BY mid ORDER BY mid"
								'Response.write strSQL3 & "<br>"
								
								oRec3.open strSQL3, oConn, 3 
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
									        
									        if oRec("usejoborakt_tp") <> 0 then '** job eller akt base. 1 = akt
									        strSQL4 = "SELECT sum(timer) AS sumfaktimer, t.kurs, t.valuta, v.valutakode, a.aktbudget FROM timer t "_
									        &" LEFT JOIN valutaer v ON (v.id = 1) "_
									        &" LEFT JOIN aktiviteter AS a ON (a.id = t.taktivitetid) "_
									        &" WHERE tjobnr = "& oRec("jobnr") &" AND tmnr = "& oRec3("mid") &""_
									        &" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' "_
									        &" AND ("& aty_sql_realHoursFakbar &") GROUP BY t.taktivitetid"
									        else
 									        strSQL4 = "SELECT sum(timer) AS sumfaktimer, t.kurs, t.valuta, v.valutakode FROM timer t "_
									        &" LEFT JOIN valutaer v ON (v.id = 1) "_
									        &" WHERE tjobnr = "& oRec("jobnr") &" AND tmnr = "& oRec3("mid") &""_
									        &" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND ("& aty_sql_realHoursFakbar &") GROUP BY timepris"
									        end if 
									else
									strSQL4 = "SELECT sum(timer) AS sumfaktimer, t.timepris, t.kurs, t.valuta, v.valutakode "_
									&" FROM timer t LEFT JOIN valutaer v ON (v.id = 1) WHERE tjobnr = "& oRec("jobnr") &" AND tmnr = "& oRec3("mid") &""_
									&" AND tdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' AND ("& aty_sql_realHoursFakbar &") GROUP BY timepris" 
									end if
									
									'Response.Write strSQL4 & "<hr>"
									'Response.flush
									
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
										
										    if oRec("usejoborakt_tp") <> 0 then '*** Job eller akt
										    faktimertimepris = oRec4("aktbudget")
										    else
										    faktimertimepris = (oRec("jobTpris") / oRec("budgettimer"))
										    'Response.write "# "& oRec("jobTpris") & " - "& oRec("budgettimer")
										    end if
										else
										thissumfaktimer = 0
										faktimertimepris = oRec("jobTpris")/1
										end if
									end if
									
									strTimer = strTimer & formatnumber(thissumfaktimer,2) &" á "
									
									call beregnValuta(faktimertimepris,oRec4("kurs"),100)
							        faktimertimepris = valBelobBeregnet
									
									strTimer = strTimer & formatnumber(faktimertimepris, 2) & " "& oRec4("valutakode") &"<br>"
									
									totsumfaktimerprmed = totsumfaktimerprmed + thissumfaktimer
									totsumfakkrprmed = totsumfakkrprmed + (thissumfaktimer * faktimertimepris) 
									
									'*jobtotaler **
									totsumfaktimer = totsumfaktimer + thissumfaktimer
									totsumfakkr = totsumfakkr + (thissumfaktimer * faktimertimepris)
									oRec4.movenext
									wend 
									oRec4.close
									
									strTimer = strTimer & "Total: "& formatnumber(totsumfaktimerprmed, 2)&" timer <br> "& formatnumber(totsumfakkrprmed, 2) &" DKK</td>"
									
									
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
									
									
									
									
									bel = 0
									
									strSQL4 = "SELECT jobid, faknr, fid, faktype, valuta, kurs FROM fakturaer"_
									&" WHERE jobid = "& usejid &" AND ((faktype = 0 OR faktype = 1) AND shadowcopy <> 1) AND fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"' ORDER BY fid " 
									
									'Response.Write strSQL4
									'Response.flush
									
									oRec4.open strSQL4, oConn, 3 
									
									while not oRec4.EOF 
										
										strFak = strFak &"<tr><td width=50 valign=top class=lille style='padding-left:2px;'><font color='#2c962d'>"& oRec4("faknr") &"</font></td>"
										
										strSQL2 = "SELECT sum(fak) AS medsumfaktimer, sum(enhedspris*fak) AS medfakbelob, kurs, valuta FROM fak_med_spec"_
										&" WHERE fakid = "& oRec4("fid") &" AND mid = " & usemid &" AND fak <> 0 AND enhedsang = 0 GROUP BY enhedspris, kurs, valuta"
										oRec2.open strSQL2, oConn, 3 
										
										strFak = strFak & "<td class=lille valign=top width=130><font color='#5582d2'>"
										while not oRec2.EOF 
											
											
											
											if len(oRec2("medsumfaktimer")) <> 0 then
											faktureredetimer = oRec2("medsumfaktimer")
											else
											faktureredetimer = 0
											end if
											
											if len(oRec2("medfakbelob")) <> 0 then
											medfakbelob = oRec2("medfakbelob")
											else
											medfakbelob = 0
											end if
											
											call beregnValuta(medfakbelob,oRec2("kurs"),100)
							                fakenhpris = valBelobBeregnet
											 
											
											'**** Kreditnota ***'
											if oRec4("faktype") = 1 then
											fakenhpris = -(fakenhpris)
											faktureredetimer = -(faktureredetimer)
											end if
											
											strFak = strFak & formatnumber(faktureredetimer, 2) &" á "& formatnumber(fakenhpris/faktureredetimer, 2) &" DKK<br>"
											faktureredetimerialt = faktureredetimerialt + (faktureredetimer)
											
											bel = bel + fakenhpris' bel + (faktureredetimer * fakenhpris)
										oRec2.movenext
										wend
										oRec2.close
										
										strFak = strFak & "</td>"
									
									
									
										    if len(bel) <> 0 then
											bel = bel 
											else
											bel = 0
											end if
											
											
									        '**** Kreditnota ***'
											if oRec4("faktype") = 1 then
											bel = -(bel)
											end if
									
									strFak = strFak & "<td width=60 class=lille align=right valign=top style='padding-right:2;'>"& formatnumber(faktureredetimerialt) &" t.</td>"
									strFak = strFak & "<td width=100 class=lille align=right valign=top style='padding-right:2;'>"& formatnumber(bel, 2) & " DKK</td></tr>"
									
									
									'**** faktotaler på medareb ***
									faktimertotalpamedpafak = faktimertotalpamedpafak + faktureredetimerialt
									fakkrprmedprfak = fakkrprmedprfak + bel
									
									'*** faktotaler på job ***
									totbelsamletfak = totbelsamletfak + bel
									totfaktim = totfaktim + faktureredetimerialt
									
									faktureredetimerialt = 0
									bel = 0
									
									oRec4.movenext
									wend
									oRec4.close 
									
									
									'*** faktureret totaler på medarbejder ***
									strFak = strFak & "<tr>"
									strFak = strFak & "<td colspan=2 width=180 align=right class=lille style='padding-right:2px;'>Total:</td>"
									strFak = strFak & "<td align=right width=60 class=lille style='padding-right:2px;'>"&formatnumber(faktimertotalpamedpafak)&" t.</td>"
									strFak = strFak & "<td align=right width=100 class=lille style='padding-right:2px;'>"&formatnumber(fakkrprmedprfak, 2)&" DKK</td>"
									strFak = strFak & "</tr>"
									strFak = strFak & "</table>"
									
									strTimer2 = "</td>"
									strTimer2 = strTimer2 & "</tr>"
								
								faktimertotalpamedpafak = 0
								fakkrprmedprfak = 0
								'*************************************************
								
								'*** Udskriver Timer og Faktimer på medarbjeder hvis de findes ***
									if thissumtimer <> 0 OR faktureredetimer <> 0 then
										
										if oRec("jid") <> lastjid then 	
										Response.Write strthisJob
										lastjid = oRec("jid")
										end if
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
									strSQL4 = "SELECT beloeb AS totbeloebpaafak, faktype, kurs, valuta, fid FROM fakturaer"_
									&" WHERE jobid = "& usejid &" AND ((faktype = 0 OR faktype = 1) AND shadowcopy <> 1) AND fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"'"_
									&" ORDER BY fid " 'GROUP BY faktype
									
									'Response.write strSQL4
									'Response.flush
									
									
									
									oRec4.open strSQL4, oConn, 3 
									while not oRec4.EOF 
									    
									    call beregnValuta(oRec4("totbeloebpaafak"),oRec4("kurs"),100)
							          
									
									    if oRec4("faktype") <> 1 then
										totbelpaafak = totbelpaafak + valBelobBeregnet
										else
										totbelpaafak = totbelpaafak -(valBelobBeregnet)
										end if
										
										
										'*** Ikke aktiv mere efter KM er i brug. Æbler og pærer ***'
										'if oRec4("faktype") <> 1 then
										'tottimerpaafak = tottimerpaafak + oRec4("tottimerpaafak")
										'else
										'tottimerpaafak = tottimerpaafak -(oRec4("tottimerpaafak"))
										'end if
										    
										    
										    
										
										
										
										'**** Materialer udsp. ***
										strSQL5 = "SELECT sum(fms.matbeloeb * ("& replace(oRec4("kurs"), ",", ".") &"/100)) AS fmsbel, "_
										&" fms.matfakid FROM fak_mat_spec fms "_
										&"WHERE (fms.matfakid = "& oRec4("fid") &") GROUP BY fms.matfakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec4("faktype")
										case 0 '*** Faktura
										fakmatBeloeb = fakmatBeloeb + oRec3("fmsbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakmatBeloeb = fakmatBeloeb - oRec3("fmsbel")
										end select
										
										
										end if
										oRec3.close
										
										
										'**** Akt sum udsp. excl. kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec4("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec4("fid") &" AND showonfak = 1 AND enhedsang <> 3) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec4("faktype")
										case 0 '*** Faktura
										fakaktBeloeb = fakaktBeloeb + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktBeloeb = fakaktBeloeb - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close
										
										
										'**** Akt sum udsp. KUN kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec4("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec4("fid") &" AND showonfak = 1 AND enhedsang = 3) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktKorsBeloeb = fakaktKorsBeloeb + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktKorsBeloeb = fakaktKorsBeloeb - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close
										    
										    
										
									
									oRec4.movenext
									wend
									oRec4.close
									
									
									  
									
									
									'if len(tottimerpaafak) <> 0 then
									'tottimerpaafak = tottimerpaafak
									'else
									'tottimerpaafak = 0
									'end if
									
									if len(totbelpaafak) <> 0 then
									totbelpaafak = totbelpaafak
									else
									totbelpaafak = 0
									end if
								 %>
								
								<%if totbelpaafak <> 0 OR totsumtimer <> 0 then%>
								<tr bgcolor="#ffffe1">	
									<td bgcolor="#e4e4e4" colspan=2 valign=top style="padding-left:5px; border:1px #8caae6 solid;" valign=bottom width=150 height=55>
									<font class=megetlillesort>
									<b><%=oRec("jobnavn")%> </b>(<%=oRec("jobnr")%>)<br />
									Jobtype: <%=jtype%><br>
									<%=strJob%>
									</td>
									<td align=right style="padding-right:2px; padding-left:2px; border:1px #8caae6 solid;" valign=top class=lille><u><%=formatnumber(totsumtimer, 2)%> t.</u></td>
									<td style="padding-left:2px; border:1px #8caae6 solid;" valign=top class=lille><u><%=formatnumber(totsumfaktimer, 2)%> t.</u> - <u><%=formatnumber(totsumfakkr, 2)%> DKK</u></td>
									<td valign=top style="border:1px #8caae6 solid;">
									<table cellspacing=0 cellpadding=0 border=0>
									<tr>
										<td valign=bottom style="padding-left:2px; width:180px;" class=lille>
										<br /><br />
										Total beløb på fakturaer på dette job:<br />
										<font color="#2c962d">
										<b><%=formatnumber(totbelpaafak, 2)%></b> DKK</font>
										
										<br />
										Aktiviteter, incl. timer, enh. og stk, samt "herreløse" aktiviteter (ekstra sum-aktiviteten): <b><%=formatnumber(fakaktBeloeb, 2) %> DKK</b><br />
										Materialer / udl.: <b> <%=formatnumber(fakmatBeloeb, 2) %> DKK</b><br />
										KM.: <b><%=formatnumber(fakaktKorsBeloeb, 2) %> DKK</b><br />
										
										</td>
										<td valign=top align=right style="padding-right:2; width:60px;" class=lille><u><%=formatnumber(totfaktim)%> t.</u></td>
										<td valign=top align=right style="padding-right:2; width:100px;" class=lille><u><%=formatnumber(totbelsamletfak)%> DKK</u></td>
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
					
					
					'lastjid = oRec("jid")
					
					
			
			
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
		
		if print <> "j" then

        ptop = 72
        pleft = 810
        pwdt = 140

        call eksportogprint(ptop,pleft,pwdt)
        %>
        <tr>
            
            <td align=center>
            <a href="erp_job_afstem.asp?print=j&fm_kunder=<%=selKunde%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&showall=<%=showall %>" target="_blank">
           &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
            </td><td><a href="erp_job_afstem.asp?print=j&fm_kunder=<%=selKunde%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&showall=<%=showall %>" class='rmenu' target="_blank">Print version</a></td>
           </tr>
           
	       
           </table>
        </div>
        <%else%>

        <% 
        Response.Write("<script language=""JavaScript"">window.print();</script>")
        %>
        <%end if
		
		end if '** Over 95 dage valgt!%>
		</div>	
<%
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
