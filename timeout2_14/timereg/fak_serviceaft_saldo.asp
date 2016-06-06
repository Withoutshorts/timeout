<%Response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="inc/dato2.asp"-->

<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else



%>
<script>

function setaftid(){
kundeid = document.getElementById("kundeid").value 
//document.getElementById("aftid").value = -1
window.location.href = "fak_serviceaft_saldo.asp?aftid=-1&kundeid="+kundeid
}

</script>

<%

if len(request("id")) <> 0 then
id = request("id")
else
id = 0
end if

thisfile = "fak_serviceaft_saldo.asp"
func = request("func")
FM_soeg = request("FM_soeg")
print = request("print")


if len(request("aftid")) then
aftid = request("aftid")
else
aftid = -1
end if

if len(request("kundeid")) <> 0 then
kundeid = request("kundeid")
else
        if len(request.cookies("erp")("kid")) <> 0 then
	    kundeid = request.cookies("erp")("kid")
		else
		kundeid = 0
		end if
end if

response.Cookies("erp")("kid") = kundeid
response.Cookies("erp").expires = date + 60


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



%>
		
		<%if print <> "j" then
		topS = 132
		leftS = 20
		else
		topS = 60
		leftS = 20
		%>
		<table cellspacing="0" cellpadding="0" border="0" width="880">
				<tr>
					<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
					<td bgcolor="#FFFFFF" align=right><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
				</tr>
				</table>
		<%
		end if
		%>
		
		<div id="sindhold" style="position:absolute; left:<%=leftS%>; top:<%=topS%>; visibility:visible;">
		<h3>Afstemning Aftaler</h3> 
				
				
				<%
				
				
				if print <> "j" then%>
				<div id="filterdiv" style="position:relative; left:0px; top:0px; visibility:visible; background-color:#ffffe1; border:1px #C4C4C4 solid; padding:10px 10px 10px 10px;">
		
				<table border=0 cellspacing=0 cellpadding=0>
				<form id=filter name=filter method=post action="<%=thisfile%>">
				<tr><td align=right style="padding:4px 5px 5px 5px;" valign=top><b>Kontakt:</b></td>
				<td valign=top style="padding:1px 5px 5px 5px;">
				<select name="kundeid" id="kundeid" style="font-size : 9px; width:285px;" onChange="setaftid()">
				<option value="0">(ingen)</option>
		
				<%
				strSQL = "SELECT s.kundeid, Kkundenavn, Kkundenr, Kid FROM serviceaft s "_
				&"LEFT JOIN kunder ON (kid = s.kundeid) WHERE s.status = 1 GROUP BY s.kundeid ORDER BY Kkundenavn"
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("kkundenr") %>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				
			</select>
			</td>
			<%if kundeid = 0 then %>
			<td>&nbsp;&nbsp;<input type="submit" value="Vælg kontakt"></td>
			<%end if %>
			</tr>
			
				
				<%if kundeid <> 0 then %>
				<tr>
				<td style="padding:1px 5px 5px 5px;" valign=top align=right><b>Aftale:</b></td>
				<td style="padding:1px 5px 5px 5px;"><select name="aftid" id="aftid" style="font-size : 9px; width:285px;">
				<option value="-1">(ingen) - Vælg aftale</option>
				
		
				<%
				strSQL = "SELECT s.id, s.enheder, s.stdato, s.sldato, s.status, s.navn, s.pris AS aftpris, s.perafg, "_
				&" s.advitype, s.advihvor, s.erfornyet, s.erfaktureret, s.varenr, "_
				&" s.aftalenr, s.fordel, s.kundeid "_
				&" FROM serviceaft s"_
				&" WHERE s.kundeid = "& kundeid &" ORDER BY s.navn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(aftid) = cint(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%> (<%=oRec("aftalenr")%>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				
			</select>
                    <input name="FM_usedatokri" value="1" type="hidden" />
			</td></tr>
				
				
				<tr>
					<td style="padding:5px;" colspan=2>
					<b>Periode:</b><br>
					<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
					&nbsp;&nbsp;<input type="submit" value="Vis Aftale"></td>
				</tr>
				<tr>
					<td colspan=2 align=right style="padding-right:30px;"><br>
					
					</td>
				</tr>
				<%else %>
                    <input id="aftid" name="aftid" value="-1" type="hidden" />
				<%end if 'kundeid %>
				
				</form>
				</table>
				</div>
				
				<%if cint(kundeid) > 0 AND cint(aftid) > 0 then %>
				<div style="position:absolute; left:325px; top:195px;"><a href="<%=thisfile%>?menu=kon&aftid=<%=aftid%>&kundeid=<%=kundeid%>&print=j" target="_blank" class=vmenu>Print venlig version&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a></div>
				<%end if %>
				
				<%else 'print
				strMrd = Request.Cookies("datoer")("st_md")
				strDag = Request.Cookies("datoer")("st_dag")
				strAar = Request.Cookies("datoer")("st_aar") 
				strDag_slut = Request.Cookies("datoer")("sl_dag")
				strMrd_slut = Request.Cookies("datoer")("sl_md")
				strAar_slut = Request.Cookies("datoer")("sl_aar")
				end if%>
		
		
<!--<font color=crimson>BETA TEST VERSION (udvalgte brugere)</font>-->




<%if print = "j" then%>
<b>Periode:</b>
<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>
<br>
<br>
<%end if


if aftid <> "-1" then


sub faktr
saldo = saldo + oRec3("enheder")
%>
<table>
<tr bgcolor="#ffffe1">
	<td><%=oRec3("fakdato")%></td>
	<td><b>Faktura:</b>&nbsp;<%=oRec3("faknr")%></td>
	<td align=right><%=formatnumber(oRec3("enheder"), 2)%></td>
	<td>&nbsp;</td>
	<td align=right><%=formatnumber(saldo, 2)%></td>
</tr>
</table>
<%
end sub

'*** SQL datoer ***
strStartDato = strAar&"/"&strMrd&"/"&strDag
strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut

if fmudato = 1 then
sqlDatoKriFak = " AND f.fakdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
sqlDatoKriFordel = " AND fordeldato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
else
sqlDatoKriFak = ""
sqlDatoKriFordel = ""
end if



strStartDatoDiff = strDag&"/"&strMrd&"/"&strAar
strSlutDatoDiff = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut

valgtperiode_antalmaneder = datediff("m", strStartDatoDiff, strSlutDatoDiff, 2, 3)

aftKri = "s.id = "& aftid &"  AND "


		strSQL = "SELECT s.id, s.enheder, s.stdato, s.sldato, s.status, s.navn, s.pris AS aftpris, s.perafg, "_
		&" s.advitype, s.advihvor, s.erfornyet, s.erfaktureret, s.varenr, aftalenr, "_
		&" k.kkundenavn, k.kkundenr, kid, k.adresse, k.postnr, k.land, k.city, "_
		&" s.aftalenr, s.fordel, s.kundeid, s.overfortsaldo, s.advitype "_
		&" FROM serviceaft s"_
		&" LEFT JOIN kunder k ON (kid = kundeid) "_
		&" WHERE "& aftKri &" s.kundeid = "& kundeid &" GROUP BY s.id ORDER BY s.id"
		'&" WHERE s.id <> 0 AND s.kundeid <> 0 GROUP BY s.id ORDER BY s.id"
		
		
		'Response.write strSQL
			
		t = 0	
		saldo = 0
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		
			'call tdbgcol_to_1(t)
			%>
			<br />
			<table cellpadding=0 cellspacing=0 border=0 width=300><tr>
			<td style="background-color:#EFF3FF; border:1px #8cAAE6 solid; padding:10px 10px 10px 10px;">
			
			
			<%if print = "j" then%>
			<b>Kunde:</b>
			<%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)<br>
			<%=oRec("adresse")%><br>
			<%=oRec("postnr")%>, <%=city%><br>
			<%=oRec("land")%><br><br>
			<%end if%>
			
			<b>Aftale:</b> <%=oRec("navn")%> (<%=oRec("aftalenr") %>)<br>
			<b>Periode:</b> <%=formatdatetime(oRec("stdato"), 1)%> - <%=formatdatetime(oRec("sldato"), 1)%><br>
			<b>Enheder tildelt:</b> <%=formatnumber(oRec("enheder"), 2)%><br>
			<b>Overført saldo:</b> <%=formatnumber(oRec("overfortsaldo"), 2)%><br>
			<b>Aftale type:</b>
			<%if oRec("advitype") = 0 then%>
			Periode
			<%else%>
			Enh. / Klip
			<%end if%><br />
			
			<b>Pris:</b> <%=formatcurrency(oRec("aftpris"), 2)%><br />
			
			<b>Job:</b><br />
			
			<%
			
			strSQLj = "SELECT id, jobnavn, jobnr FROM job WHERE serviceaft = " & oRec("id")
			oRec2.open strSQLj, oConn, 3
		    while not oRec2.EOF
		    %>
		    <%=oRec2("jobnavn") &" ("& oRec2("jobnr")&")" %><br />
		    <%
		    
		    oRec2.movenext
		    wend
		    oRec2.close 
			
			
			%>
			
			</td>
			</tr></table>			
						
			
			<%
			
			overfortsaldo = oRec("overfortsaldo")
			aftStDato = oRec("stdato")
			aftSlDato = oRec("sldato")
			aftid = oRec("id")
			intFordeling = oRec("fordel")
			aftaleMdPerDiff = datediff("m", oRec("stdato"), oRec("sldato"), 2, 3) + 1
			intEnheder = oRec("enheder")
			
			'Response.write aftaleMdPerDiff
			
			if intFordeling = 1 then
			enhederPrMd = (intEnheder/aftaleMdPerDiff)
			else
			enhederPrMd = intEnheder
			end if
			
			aftpris = oRec("aftpris")
			aftprisPrMd = (aftpris/aftaleMdPerDiff)
			
			
		end if
		oRec.close
		
		if enhederPrMd <> 0 then
		enhederPrMd = enhederPrMd
		else
		enhederPrMd = 0
		end if
		
		if aftprisPrMd <> 0 then
		aftprisPrMd = aftprisPrMd
		else
		aftprisPrMd = 0
		end if
		
		FakenhederAkk = 0
		
		
			
			
			
			
			%>
			<br />
			<table cellspacing=1 cellpadding=0 border=0 width=95%>
			<tr>
			    <td bgcolor="#FFFFFF" align=right style="padding:0px 5px 0px 0px;"><b>Måneder</b></td>
				<td colspan="1" bgcolor="#FFFFe1" align=center style="padding:0px 2px 0px 2px;"><b>Aftale pris</b></td>
				<td colspan="6" align=center bgcolor="#EFF3FF"><b>Aftale enheder tildelt / realiseret</b> </td>
				<td colspan="5" bgcolor="yellowgreen" align=center><b>Aftale enheder faktureret</b></td>
				<td colspan="1" bgcolor="yellowgreen" align=center style="padding:0px 2px 0px 2px;"><b>Faktureret kr.</b></td>
				<td colspan="2" align=center bgcolor="#FFFFe1"><b>Saldo kr.</b><br></td>
			    <td colspan="4" align=center bgcolor="#FFFFFF"><b>Grundlag<br />timer og materialer</b> </td>
			</tr>
			<!--<tr>
			
				<td align=left width=80 bgcolor="#FFFFFF">
                    &nbsp;</td>
				<td align=right class=lillegray align=center bgcolor="#EFF3FF">Enheder</td>
				<td align=right bgcolor="#EFF3FF">Enheder</td>
				<td align=right bgcolor="#EFF3FF">Enheder</td>
				<td align=right bgcolor="#EFF3FF">Enheder</td>
				<td align=right bgcolor="#EFF3FF">Saldo Enh.</td>
				<td align=right bgcolor="#EFF3FF">Saldo Enh.</td>
				<td align=right>Enheder</td>
				<td align=right>Enheder</td>
				<td align=right>Saldo Enh.</td>
				<td align=right>Saldo Enh.</td>
				<td align=right>Saldo Enh.</td>
				<td align=right>Kr.</td>
				<td align=right>Kr.</td>
				<td align=right>Saldo Kr.</td>
				<td class=right align=right>Saldo Kr.</td>
			</tr>-->
			
			<tr>
				<td align=right style="padding:0px 5px 0px 0px;" bgcolor="#FFFFFF">
                    &nbsp;</td>
				<td valign=bottom bgcolor="#FFFFe1" align=right style="padding:0px 2px 0px 2px;">Aftale<br> beløb fordelt pr. md.</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Tildelt på aftale</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Overført saldo</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Real. enheder<br>(timer * faktor)</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Real. Akku.</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Tildelt / Real.</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Tildelt / Real. Akku.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Fakt. antal enh.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Fakt. Akku.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Tildelt på aft. / Fakt.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Tildelt / Fakt. Akku.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Real. / Fakt. Akku.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Faktureret</td>
				<td valign=bottom align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">Aftale beløb pr. md / Fak. pr. md</td>
				<td valign=bottom align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">Aftale beløb / Faktureret Akku.</td>
			    <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">Real. timer antal</td>
				<td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">Realiserede timer beløb</td>
				<td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">Real. mat. antal</td>
				<td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">Real. materialer beløb</td>
				
			</tr>
			
			
			<%
			lastyear = 0
			tilfojSaldo = 0
			
			For x = 0 to valgtperiode_antalmaneder
			
						
			'*** Start år og md ***
			if x = 0 then
			useyear = datepart("yyyy", strStartDatoDiff)
			usemd = datepart("m", strStartDatoDiff)
			usedag = 1
			
			else
			
			usemd = datepart("m", dateadd("m", x, strStartDatoDiff))
			nyDato = dateadd("m", x, strStartDatoDiff)
			useyear = datepart("yyyy", nyDato)
			usedag = 1
			
			end if
			
			nyPerDato = usedag&"/"&usemd&"/"&useyear
			nextPerDato = dateadd("m", 1, nyPerDato)
			nyPerStDatoSQL = useyear&"/"&usemd&"/"&usedag
			nyPerSlDatoSQL_temp = dateadd("m", 1, nyPerDato)
			nyPerSlDatoSQL = year(nyPerSlDatoSQL_temp)&"/"&month(nyPerSlDatoSQL_temp)&"/"&day(nyPerSlDatoSQL_temp)
			
			if lastyear <> useyear then
			call tdbgcol_to_1(x)%>
				<tr bgcolor="#c4c4c4" height=20>
					<td colspan=20 align=center><b><%=useyear%></b></td>
			    </tr>
			<%
			end if
			
			lastyear = useyear
			
			
						'*** Aftale 
						'saldo = saldo - oRec("enheder")
						call tdbgcol_to_1(x)%>
						<tr>
							<td bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><b><%=monthname(usemd)%>&nbsp;<%=useyear%></b></td>
							
							<td align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">
							<%
							'*** Aft. pris pr. md. ***
							if (datepart("m",nyPerDato, 2,3) >= datepart("m",aftStDato, 2,3) AND datepart("yyyy",nyPerDato, 2,3) >= datepart("yyyy",aftStDato, 2,3)) AND cDate(aftSlDato) >= cDate(nyPerDato) then
							Response.write formatcurrency(aftprisPrMd, 2)
							usePrisTilSaldoprMd = aftprisPrMd 
							else
							Response.write formatcurrency(0, 2)
							usePrisTilSaldoprMd = 0
							end if%>
							
							<!--<br /><=nyPerDato %> <br /> <=aftStDato%>-->
							</td>
							
							<td bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<%
							'*** Hvis manuel månedsfordeling er valgt ****
							if intFordeling = 1 then
									
											'**** Fordeling ***
											strSQL2 = "SELECT enheder FROM aft_enh_fordeling WHERE aft_id ="& aftid &" AND maned = "& usemd & " AND aar = "& useyear 
											oRec2.open strSQL2, oConn, 3
											
											'response.write strSQL2
											if not oRec2.EOF then
											enhederPrMd = oRec2("enheder")
												
												'*** Tilføj overført saldo ***
												if tilfojSaldo = 0 then
												tilfojSaldo = 1
												else
												tilfojSaldo = 2
												end if
											
											else
											enhederPrMd = 0
											end if
											oRec2.close
									
									
									
									
									Response.write formatnumber(enhederPrMd, 2)
									useEnhTilSaldo = enhederPrMd
									
							else
							       '*** Hvis måned er indenfor aftaleperiode
								 
								   '*** Er aftale per / klip afhængig???
								  'Response.write cDate(nyPerDato) &"<="& cDate(aftStDato) &"AND"& cDate(aftStDato) &"<="& cDate(nextPerDato) & "<br>"
								   if cDate(nyPerDato) <= cDate(aftStDato) AND cDate(aftStDato) <= cDate(nextPerDato) then
								   	'Response.write "perOK!<br>"
								  	 Response.write formatnumber(enhederPrMd, 2)
									 useEnhTilSaldo = enhederPrMd
									 tilfojSaldo = 1
								   else
								   	'enhederPrMd = 0
								   	Response.write formatnumber(0, 2)
									useEnhTilSaldo = 0
								   end if
								    
								   
							end if
							%>
							</td>
							
							<td bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(overfortsaldo, 2) %>
							<%overfortsaldo = 0 %>
							</td>
							
							<td bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<%
							'***** Timer / enheder Realiseret i måned ***
							enhederRealiseret = 0
							timerThis = 0
							timerOms = 0
							strSQLuds = "SELECT j.id, j.jobnr, j.jobnavn, j.jobnr, t.timer, t.tdato, "_
							&" t.taktivitetid, a.faktor, t.Tmnavn, t.Tmnr, j.jobstartdato, t.timepris "_
							&" FROM timer t "_
							&" LEFT JOIN job j ON (j.jobnr = t.tjobnr "& jobkundeSQLkri &")"_
							&" LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) "_
							&" WHERE t.seraft = "& aftid &" AND tfaktim = 1 AND tdato BETWEEN '"& nyPerStDatoSQL &"' AND '"& nyPerSlDatoSQL &"'"
							
							oRec.open strSQLuds, oConn, 3
							while not oRec.EOF
							
							enhederRealiseret = enhederRealiseret + (oRec("faktor") * oRec("timer"))
							timerThis = timerThis + oRec("timer")
							timerOms = timerOms + (oRec("timer") * oRec("timepris"))
							
							oRec.movenext
							wend
							oRec.close
							
							if len(enhederRealiseret) <> 0 then
							enhederRealiseret = enhederRealiseret
							else
							enhederRealiseret = 0
							end if
							
							'Response.write tilfojSaldo 
							
							if tilfojSaldo = 1 then
							enhederRealiseret = enhederRealiseret + (overfortsaldo)
							tilfojSaldo = 2
							star = "*"
							end if
							
							Response.write formatnumber(enhederRealiseret, 2) & star 
							
							star = ""
							
							%>
							</td>
							
							<td bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<%
							enhederRealiseretAkk = (enhederRealiseretAkk + enhederRealiseret)
							Response.write formatnumber(enhederRealiseretAkk, 2)
							%>
							</td>
							
							<td bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<%
							saldomd = (useEnhTilSaldo - enhederRealiseret)
							saldoAkk = (saldoAkk +(saldomd)) 
							Response.write formatnumber(saldomd, 2)
							%>
							</td>
							
							<td bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<b><%=formatnumber(saldoAkk, 2)%></b>
							</td>
							
							
							
							<td align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
								
								<%
								'**** Fakturaer i måned ***
									
										strSQL3 = "SELECT f.fakdato, f.fid, sum(f.timer) AS enheder, sum(f.beloeb) AS fakbelob, f.faknr FROM fakturaer f WHERE f.fakdato BETWEEN '"& nyPerStDatoSQL &"' AND '"& nyPerSlDatoSQL &"' AND f.aftaleid = "& aftid &" GROUP BY f.aftaleid"
										oRec3.open strSQL3, oConn, 3
										
										FakKrprMd = 0
										FakenhederPrMd = 0
										if not oRec3.EOF then
										FakKrprMd = oRec3("fakbelob")
										FakenhederPrMd = oRec3("enheder")
										end if
										
										oRec3.close
								
								Response.write formatnumber(FakenhederPrMd, 2)
								FakenhederAkk = (FakenhederAkk +(FakenhederPrMd))%>
							
							</td>
							<td align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;"><%=formatnumber(FakenhederAkk, 2)%></td>
							<td align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_Fak = (FakenhederPrMd - useEnhTilSaldo)
							Response.write formatnumber(enhSaldoTild_Fak, 2) 
							%>
							</td>
							<td align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_FakAkk = (enhSaldoTild_FakAkk + (enhSaldoTild_Fak))
							Response.write formatnumber(enhSaldoTild_FakAkk, 2) 
							%>
							</td>
							<td align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoReal_Fak = (enhederRealiseret - FakenhederPrMd)
							enhSaldoReal_FakAkk = (enhSaldoReal_FakAkk + (enhSaldoReal_Fak))
							  
							%>
							<b><%=formatnumber(enhSaldoReal_FakAkk, 2)%></b>
							</td>
							
							
							
							
							<td align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;"><%=formatcurrency(FakKrprMd, 2)%></td>
							
							
							<td align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_Fak = (FakKrprMd - usePrisTilSaldoprMd)
							Response.write formatcurrency(krSaldopris_Fak, 2) 
							%>
							</td>
							<td align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_FakAkk = (krSaldopris_FakAkk + (krSaldopris_Fak))
							%>
							<b><%=formatcurrency(krSaldopris_FakAkk, 2) %></b>
							</td>
							
							<td bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(timerThis, 2)%></td>
							<td bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatcurrency(timerOms, 2)%></td>
							
							
							
							<% 
							'***** Materialer Realiseret i måned ***
							matRealiseret = 0
							matPris = 0
							
							 strSQLmat = "SELECT sum(matantal) AS matAntal, sum(matsalgspris) AS matPris FROM materiale_forbrug WHERE "_
                            &" serviceaft = " & aftid & ""_ 
	                        &" AND forbrugsdato BETWEEN '"& nyPerStDatoSQL &"' AND '"& nyPerSlDatoSQL  &"' GROUP BY serviceaft"
	                         
							
							oRec.open strSQLmat, oConn, 3
							if not oRec.EOF then
							
							matRealiseret = oRec("matAntal")
							matPris = oRec("matPris")
							
							
							end if
							oRec.close
							%>
							<td bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(matRealiseret, 2)%></td>
							<td bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatcurrency(matPris)%></td>
							
							
						</tr>
						<%
next

%>

				
		
</table>
<br>
* = Enheder realiseret incl. evt. overført saldo.

<%end if  'aftid%>


		</div>
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->