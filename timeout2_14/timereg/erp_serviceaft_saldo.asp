<%Response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else



%>


<%

if len(request("id")) <> 0 then
id = request("id")
else
id = 0
end if


menu = "erp"
thisfile = "erp_serviceaft_saldo.asp"
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
<script>

function setaftid(){
kundeid = document.getElementById("kundeid").value 
//document.getElementById("aftid").value = -1
window.location.href = "erp_serviceaft_saldo.asp?aftid=-1&kundeid="+kundeid
}

</script>



<%
    
 call menu_2014()   
    
 else%>

<% 
Response.Write("<script>Javascript:window.print()</script>")
%>

<!--#include file="../inc/regular/header_hvd_inc.asp"-->


<%end if



%>
		
		<%if print <> "j" then
		topS = 102
		leftS = 90
		else
		topS = 20
		leftS = 20
		
		end if
		%>
		
		<div id="sindhold" style="position:absolute; left:<%=leftS%>px; top:<%=topS%>px; visibility:visible;">
        <% 


            'oimg = "view_1_1.gif"
	oleft = 0
	otop = 0
	owdt = 700
	oskrift = "Afstemning aftaler"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
    
		
		
	            if print <> "j" then

                call filterheader_2013(0,0,800,pTxt)
        
       
				
				
				%>
				
				<table border=0 cellspacing=0 cellpadding=0 width=100%>
				<form id=filter name=filter method=post action="<%=thisfile%>">
				<tr><td align=right style="padding:10px 5px 5px 5px;" valign=top><b>Kontakt:</b></td>
				<td valign=top style="padding:6px 5px 5px 5px;">
				
				<%
				strSQL = "SELECT s.kundeid, Kkundenavn, Kkundenr, Kid FROM serviceaft s "_
				&"LEFT JOIN kunder ON (kid = s.kundeid) WHERE s.status = 1 GROUP BY s.kundeid ORDER BY Kkundenavn"
				'Response.Write strSQL
				'Response.flush
				
				
				%>
				
				<select name="kundeid" id="kundeid" style="width:285px;" onChange="setaftid()">
				<option value="0">(ingen)</option>
		
				<%
				
				
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if len(trim(oRec("Kid"))) <> 0 then
				thisKid = oRec("Kid")
				else
			    thisKid = 0
			    end if
				
				if cint(kundeid) = cint(thisKid) then
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
			<td align=right>&nbsp;&nbsp;<input type="submit" value="Vælg kontakt >>"></td>
			<%end if %>
			</tr>
			
				
				<%if kundeid <> 0 then %>
				<tr>
				<td style="padding:1px 5px 5px 5px;" valign=top align=right><b>Aftale:</b></td>
				<td style="padding:1px 5px 5px 5px;"><select name="aftid" id="aftid" style="width:285px;">
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
                    
					<td style="padding:5px 5px 2px 72px;" colspan=2>
					<b>Periode:</b><br>
					<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
					</td>
				</tr>
				<tr>
					<td colspan=2 align=right style="padding-right:10px;">
					&nbsp;&nbsp;<input type="submit" value="Vis Aftale >>">
					</td>
				</tr>
				<%else %>
                    <input id="aftid" name="aftid" value="-1" type="hidden" />
				<%end if 'kundeid %>
				
				</form>
				</table>

				<!-- filterDiv -->
	</td></tr>
	</table>
	</div>
				
				
				
				 
				
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
			<div style="position:relative; background-color:#FFFFFF; width:1000px; padding:20px;">
			<table cellpadding=2 cellspacing=0 border=0 width=100%><tr>
			
			<tr>
			    <td colspan=2>
                    <img src="../ill/ac0009-16.gif" />&nbsp;<b>Kunde (kontakt):</b><br />
			    <%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)<br />
			    <%=oRec("adresse")%><br />
			    <%=oRec("postnr")%>, <%=city%><br />
		        <%=oRec("land")%><br />
                &nbsp;</td>
			</tr>
			
			
			<tr><td width=100><b>Aftale:</b></td><td><%=oRec("navn")%> (<%=oRec("aftalenr") %>)</td></tr>
			<tr><td><b>Periode:</b></td><td> <%=formatdatetime(oRec("stdato"), 1)%> - <%=formatdatetime(oRec("sldato"), 1)%></td></tr>
			<tr><td><b>Enheder tildelt:</b></td><td> <%=formatnumber(oRec("enheder"), 2)%></td></tr>
			<tr><td><b>Overført saldo:</b></td><td> <%=formatnumber(oRec("overfortsaldo"), 2)%></td></tr>
			
			<tr><td><b>Aftale type:</b></td><td>
			<%if oRec("advitype") = 0 then%>
			Periode
			<%else%>
			Enh. / Klip
			<%end if%></td></tr>
			
			<tr><td><b>Pris:</b></td><td> <b><%=formatcurrency(oRec("aftpris"), 2)%></b>
			<%if cint(oRec("fordel")) = 1 then %>
			(Måneds fordeling benyttet)
			<%end if %></td></tr>
			<tr><td valign=top><b>Job tilknyttet:</b></td><td>
			
			<%
			
			strSQLj = "SELECT id, jobnavn, jobnr FROM job WHERE serviceaft = " & oRec("id")
			oRec2.open strSQLj, oConn, 3
		    j = 0
		    while not oRec2.EOF
		    if j > 0 then %>
		    ,&nbsp;
		    <%end if %>
		    <%=oRec2("jobnavn") &" ("& oRec2("jobnr")&")" %>
            
		    <%
		    j = j + 1
		    oRec2.movenext
		    wend
		    oRec2.close 
			
			
			%>
			
			</td>
			</tr></table>	
			</div>		
						
			
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
			
			if intEnheder > 0 then 
			aftprisPrMd = (aftpris/intEnheder)
			else
			aftprisPrMd = aftpris
			end if
			
			
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
			<table cellspacing=1 cellpadding=0 border=0 bgcolor="#999999" width=95%>
			<tr>
			    <td bgcolor="#FFFFFF" align=right style="padding:0px 5px 0px 0px;"><b>Måneder</b></td>
				<td colspan="2" bgcolor="#FFFFe1" align=center style="padding:0px 2px 0px 2px;"><b>Aftale enh. / pris</b></td>
				<td colspan="4" align=center bgcolor="#EFF3FF"><b>Aftale enheder tildelt / realiseret</b> </td>
				<td bgcolor="yellowgreen" align=center><b>Aftale enh. faktureret</b></td>
				<td colspan="3" bgcolor="#99cc66" align=center>
                    &nbsp;</td>
				<td colspan="1" bgcolor="yellowgreen" align=center style="padding:0px 2px 0px 2px;"><b>Faktureret kr.</b></td>
				<td colspan="1" bgcolor="lightgrey" align=center style="padding:0px 2px 0px 2px;"><b>Omsætning kr.</b></td>
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
				<td valign=bottom bgcolor="#FFFFe1" align=right style="padding:0px 2px 0px 2px;">Tildelt på aftale</td>
				<td valign=bottom bgcolor="#FFFFe1" align=right style="padding:0px 2px 0px 2px;">Aftale<br> beløb fordelt pr. md.</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Overført saldo</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Real. enheder<br>(timer * faktor)</td>
				
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Tildelt / Real.</td>
				<td valign=bottom bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">Tildelt / Real. Akku.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right style="padding:0px 2px 0px 2px;">Fakt. antal enh.</td>
			    <td valign=bottom bgcolor="#99cc66" align=right style="padding:0px 2px 0px 2px;">Tildelt på aft. / Fakt.</td>
				<td valign=bottom bgcolor="#99cc66" align=right style="padding:0px 2px 0px 2px;">Tildelt på aft. / Fakt. Akku.</td>
				<td valign=bottom bgcolor="#99cc66" align=right style="padding:0px 2px 0px 2px;">Real. / Fakt. Akku.</td>
				<td valign=bottom bgcolor="yellowgreen" align=right width=80 style="padding:0px 2px 0px 2px;">Faktureret***</td>
				<td valign=bottom bgcolor="lightgrey" align=right width=80 style="padding:0px 2px 0px 2px;">Omsætning**</td>
				<td valign=bottom align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">Aftale beløb pr. md / Fak. pr. md</td>
				<td valign=bottom align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">Aftale beløb / Faktureret Akku.</td>
			    <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">Real. fak.bare timer antal</td>
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
				<tr bgcolor="#FFFFFF" height=20>
					<td colspan=19 align=center><b><%=useyear%></b></td>
			    </tr>
			<%
			end if
			
			lastyear = useyear
			
			
						'*** Aftale 
						'saldo = saldo - oRec("enheder")
						call tdbgcol_to_1(x)%>
						<tr>
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><b><%=monthname(usemd)%>&nbsp;<%=useyear%></b></td>
							
							<td class=lille bgcolor="#FFFFe1" align=right style="padding:0px 2px 0px 2px;">
							<%
							'*** Hvis manuel månedsfordeling er valgt ****
							if cint(intFordeling) = 1 then
									
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
							
							
							<td align=right class=lille bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">
							<%
							if cint(intFordeling) = 1 then
							        '*** Aft. pris pr. md. ***
							        if (datepart("m",nyPerDato, 2,3) >= datepart("m",aftStDato, 2,3) AND datepart("yyyy",nyPerDato, 2,3) >= datepart("yyyy",aftStDato, 2,3)) AND cDate(aftSlDato) >= cDate(nyPerDato) then
							        'Response.write formatcurrency(aftprisPrMd, 2)
							        Response.write formatcurrency(aftprisPrMd*enhederPrMd)
							        usePrisTilSaldoprMd = aftprisPrMd*enhederPrMd
							        else
							        Response.write formatcurrency(0, 2)
							        usePrisTilSaldoprMd = 0
							        end if
							else
							
							        if datepart("m",nyPerDato, 2,3) = datepart("m",aftStDato, 2,3) then
							        Response.Write "<b>" & formatcurrency(aftpris) & "</b>"
							        usePrisTilSaldoprMd = aftpris
							        else
        							Response.write formatcurrency(0, 2)
        							usePrisTilSaldoprMd = 0
							        end if
							        
							        
							
							
							end if%>
							
							<!--<br /><=nyPerDato %> <br /> <=aftStDato%>-->
							</td>
							
							
							
							
							
							
							
							<td class=lille bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(overfortsaldo, 2) %>
							<%overfortsaldo = 0 %>
							</td>
							
							<td class=lille bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<%
							'***** Timer / enheder Realiseret i måned ***
							enhederRealiseret = 0
							timerThis = 0
							timerOms = 0
							strSQLuds = "SELECT j.id, j.jobnr, j.jobnavn, j.jobnr, j.fastpris, j.jobTpris, j.budgettimer, t.timer, t.tdato, "_
							&" t.taktivitetid, a.faktor, t.Tmnavn, t.Tmnr, j.jobstartdato, t.timepris "_
							&" FROM timer t "_
							&" LEFT JOIN job j ON (j.jobnr = t.tjobnr "& jobkundeSQLkri &")"_
							&" LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) "_
							&" WHERE t.seraft = "& aftid &" AND tfaktim = 1 AND tdato BETWEEN '"& nyPerStDatoSQL &"' AND '"& nyPerSlDatoSQL &"'"
							
							oRec.open strSQLuds, oConn, 3
							while not oRec.EOF
							
							enhederRealiseret = enhederRealiseret + (oRec("faktor") * oRec("timer"))
							timerThis = timerThis + oRec("timer")
							'*** Fastpris / Lbn timer
							if cint(oRec("fastpris")) <> 1 then
							timerOms = timerOms + (oRec("timer") * oRec("timepris"))
							else
							timerOms = timerOms + (oRec("timer")*(oRec("jobTpris")/oRec("budgettimer")))
							end if
							
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
							
							enhederRealiseretAkk = (enhederRealiseretAkk + enhederRealiseret)
							%>
							</td>
							
							
							<td class=lille bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<%
							saldomd = (useEnhTilSaldo - enhederRealiseret)
							saldoAkk = (saldoAkk +(saldomd)) 
							Response.write formatnumber(saldomd, 2)
							%>
							</td>
							
							<td class=lille bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;">
							<b><%=formatnumber(saldoAkk, 2)%></b>
							</td>
							
							
							
							<td class=lille align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
								
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
								FakenhederAkk = (FakenhederAkk +(FakenhederPrMd))
								FakKrprMdAkk = (FakKrprMdAkk + (FakKrprMd))
								%>
							
							</td>
							
							<td class=lille align=right bgcolor="#99cc66" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_Fak = (FakenhederPrMd - useEnhTilSaldo)
							Response.write formatnumber(enhSaldoTild_Fak, 2) 
							%>
							</td>
							<td class=lille align=right bgcolor="#99cc66" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_FakAkk = (enhSaldoTild_FakAkk + (enhSaldoTild_Fak))
							Response.write formatnumber(enhSaldoTild_FakAkk, 2) 
							%>
							</td>
							
							<td class=lille align=right bgcolor="#99cc66" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoReal_Fak = (enhederRealiseret - FakenhederPrMd)
							enhSaldoReal_FakAkk = (enhSaldoReal_FakAkk + (enhSaldoReal_Fak))
							  
							%>
							<%=formatnumber(enhSaldoReal_FakAkk, 2)%>
							</td>
							
							
						
                            <td class=lille align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;"><b><%=formatcurrency(FakKrprMd, 2)%></b></td>
							<td class=lille align=right bgcolor="lightgrey" style="padding:0px 2px 0px 2px;">
							<%
							'*** Omsæting (fakturaer på job, tilknyttet aftalen)
							omsKrprMd = 0
							omsTimerPrMd = 0 
										    
								        strSQLjob = "SELECT j.id AS jobid FROM job j WHERE "_
								        &" j.serviceaft = "& aftid  
								        
								        oRec3.open strSQLjob, oConn, 3
										while not oRec3.EOF
										    
										    strSQLfakpajob = "SELECT f.fakdato, f.fid, "_
										    &" sum(f.timer) AS faktimer, sum(f.beloeb) AS fakbelob, "_
										    &" f.faknr FROM fakturaer f WHERE f.fakdato "_
										    &" BETWEEN '"& nyPerStDatoSQL &"' AND '"& nyPerSlDatoSQL &"'"_
										    &" AND f.jobid = "& oRec3("jobid") &" AND shadowcopy <> 1 GROUP BY f.jobid"
									         
									         
									       'Response.Write strSQLfakpajob
									       'Response.flush 
									       
								            oRec2.open strSQLfakpajob, oConn, 3
										    if not oRec2.EOF then
										    omsKrprMd = omsKrprMd + oRec2("fakbelob")
										    'omsTimerPrMd = oRec2("faktimer")
										    end if
										    
										    oRec2.close
										    
										    
									    oRec3.movenext
										wend
										oRec3.close
								        
                             
							   Response.Write formatcurrency(omsKrprMd)
							   omsKrprMdAkk = (omsKrprMdAkk +(omsKrprMd))
							  
							%>
							</td>
							
							
							<td class=lille align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_Fak = (FakKrprMd - usePrisTilSaldoprMd)
							Response.write formatcurrency(krSaldopris_Fak, 2) 
							%>
							</td>
							<td class=lille align=right bgcolor="#FFFFe1" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_FakAkk = (krSaldopris_FakAkk + (krSaldopris_Fak))
							%>
							<b><%=formatcurrency(krSaldopris_FakAkk, 2) %></b>
							</td>
							
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(timerThis, 2)%></td>
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatcurrency(timerOms, 2)%></td>
							
							
							
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
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(matRealiseret, 2)%></td>
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatcurrency(matPris)%></td>
							
							
						</tr>
						<%
next

%>

<tr bgcolor="#ffffff">
<td colspan=4>&nbsp;</td>
<td align=right bgcolor="#EFF3FF" class=lille style="padding:0px 2px 0px 2px;"><b><u><%=formatnumber(enhederRealiseretAkk, 2)%></u></b></td>
<td colspan="2">&nbsp;</td>
<td align=right bgcolor="yellowgreen" class=lille style="padding:0px 2px 0px 2px;"><b><u><%=formatnumber(FakenhederAkk, 2)%></u></b></td>
<td colspan="3">&nbsp;</td>
<td align=right bgcolor="yellowgreen" class=lille style="padding:0px 2px 0px 2px;"><b><u><%=formatcurrency(FakKrprMdAkk) %></u></b></td>
<td align=right bgcolor="lightgrey" class=lille style="padding:0px 2px 0px 2px;"><b><u><%=formatcurrency(omsKrprMdAkk) %></u></b></td>

    <td colspan=6>
    &nbsp;</td>
</tr>			
		
</table>
<br>


 <%
        	if print <> "j" then

        ptop = 65
        pleft = 840
        pwdt = 140

        call eksportogprint(ptop,pleft,pwdt)
        %>
       
           
	       
          




     <tr>
            
            <td align=center>
            <a href="<%=thisfile%>?menu=kon&aftid=<%=aftid%>&kundeid=<%=kundeid%>&print=j" target="_blank">
           &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
            </td><td><a href="<%=thisfile%>?menu=kon&aftid=<%=aftid%>&kundeid=<%=kundeid%>&print=j" class='rmenu' target="_blank">Print version</a></td>
           </tr>
      </table>


        </div>
        <%else%>

        <% 
        Response.Write("<script language=""JavaScript"">window.print();</script>")
        %>
        <%end if
        %>
       


       <br /><br /><br />
<div id=sidenoter style="position:relative; background-color:#FFFFFF; width:400px; border:0px red solid; padding:20px;">
<b>
<img src="../ill/ac0005-24.gif" />Side noter:</b><br />
* Enheder realiseret incl. evt. overført saldo.
<br />
** Fakturaer oprettet på job tilknyttet den valgte aftale.<br />
*** Fakturaer på aftale.<br />
Kun fakturerbare timer er medregnet i realiserede timer og enheder.<br />
Alle fakturerede beløb er uden evt. rykker beløb.
</div>
<br /><br />
            &nbsp;
<%end if  'aftid%>


		</div>
	<%end if%>
	<br /><br />
            &nbsp;
<!--#include file="../inc/regular/footer_inc.asp"-->