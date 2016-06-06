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
window.location.href = "erp_serviceaft_saldo.asp?aftid=-1&kundeid="+kundeid
}

</script>

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

    $(document).ready(function () {

        $("#FM_kunde").change(function () {

            $("#FM_sog").val('')
            $("#afstem").submit()

        });

        $("#FM_aftale").change(function () {

            $("#FM_sog").val('')
            $("#afstem").submit()

        });

        $("#FM_job").change(function () {

            $("#FM_sog").val('')
            $("#afstem").submit()

        });




    });
	
	</script>

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
<% 
Response.Write("<script>Javascript:window.print()</script>")
%>
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
		<table cellspacing="0" cellpadding="0" border="0" width="650">
				<tr>
					<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
					<!--<td bgcolor="#FFFFFF" align=right><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>-->
				</tr>
				</table>
		<%
		end if
		%>
		
		<div id="sindhold" style="position:absolute; left:<%=leftS%>; top:<%=topS%>; visibility:visible;">


       
	<h4><img src="../ill/ac0045-24.gif" /> Afstemning job og aftaler md/md</h4>
	
	<%
	
	ptop = 0
	pleft = 0
	pwdt = 802
	
	call filterheader(ptop,pleft,pwdt,pTxt) %>
	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<form action="erp_serviceaft_saldo.asp" id="afstem" method="POST">
	<tr>
	<td valign=top width=100>
	
    <% 
    if len(request("FM_kunde")) <> 0 then
			
			kundeid = request("FM_kunde")
			
	else
		
			if len(request.cookies("erp")("kid")) <> 0 AND request.cookies("erp")("kid") <> 0 then
			kundeid = request.cookies("erp")("kid")
			else
			kundeid = 0
			end if
			
	end if
	
	
	response.Cookies("erp")("kid") = kundeid
	response.Cookies("erp").expires = date + 60
	
	if len(request("FM_sog")) <> 0 AND request("FM_sog") <> "9999" then
	showSog = request("FM_sog")
	else
	showSog = ""
	end if
	
	if len(request("FM_job")) <> 0 then
	jobid = request("FM_job")
	else
	jobid = 0
	end if
	
	if len(request("FM_aftale")) <> 0 then
	aftid = request("FM_aftale")
	else
	aftid = 0
	end if
	
	if len(request("FM_viskunabne")) <> 0 AND request("FM_viskunabne") <> "0" then
	viskunabne = 1
	else
	viskunabne = 0
	end if
	
	if len(trim(showSog)) <> 0 then
	aftid = 0
	jobid = 0
	kundeid = 0
	end if
	
	select case jobid
	case -1
	vlgtJob = "Ingen"
	case 0
	vlgtJob = "Alle"
	end select 
	
	select case aftid
	case -1
	vlgtAft = "Ingen"
	case 0
	vlgtAft = "Alle"
	end select 
	
	if len(trim(request("FM_jobikkedelafaftale"))) <> 0 then
	jidaftCHK = "CHECKED"
	jidaft = 1
	else
	jidaftCHK = ""
	jidaft = 0
	end if
	
    
    %>
    <b>Kontakter:</b></td><td valign=top>
        <%if print <> "j" then %>
        <select id="FM_kunde" name="FM_kunde" style="font-size : 9px; width:325px;"> 
        
		<option value="0">Alle</option>
		<%end if
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid, COUNT(f.fid) AS antalfak FROM kunder "_
				&" LEFT JOIN fakturaer f ON (f.fakadr = kid AND shadowcopy = 0) "_
				&" WHERE "& ketypeKri &" "& strKundeKri &" AND f.fid <> 0 GROUP BY kid ORDER BY Kkundenavn"
				
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("kkundenr") %>)</option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		        </select><br>
		        <% 
		        else
		        %>
				<%=vlgtKunde %>
				<% 
				end if
				%>
	    
	    <!--
	    <b>Kontaktansvarlige (medarbejder):</b><br> <select name="FM_kans" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
	    </select>-->
	    
	    
	    
	    </td>
	    
	  
	    
	    
	  
       
       
	   
	    
	    <% if print <> "j" then%>
        <td rowspan=3 style="padding:2px 2px 2px 20px;" valign=top>
	    &nbsp;
	    
        <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
       
        </td>
	    <%else
	    dontshowDD = 1
	    %>
	    <!--#include file="inc/weekselector_s.asp"-->
	    <%end if %>
	    
	 </tr>
	 
	  <tr><td valign=top>
       
        <b>Aftaler:</b>
        </td><td valign=top>
	    
	    <% 
	    strSQLaft = "SELECT s.id AS sid, s.navn, s.aftalenr, COUNT(f.fid) AS antalfak "_
	    &" FROM serviceaft s "_
	    &" LEFT JOIN fakturaer f ON (f.aftaleid = s.id AND shadowcopy = 0) "_
	    &" WHERE s.kundeid = " & kundeid & " AND f.fid <> 0 GROUP BY s.id ORDER BY s.navn"
	  
	  'Response.Write strSQLjob
	  'Response.flush
	    %>
	    
	        <% if print <> "j" then%>
	        <select id="FM_aftale" name="FM_aftale" size=3 style="font-size : 9px; width:325px;">
            
            <%if cint(aftid) = -1 then
            ingSEL = "SELECTED"
            else
            ingSEL = ""
            end if
            %>
            
            <option value=-1 <%=ingSEL %>>Ingen</option>
            <%if cint(aftid) = 0 then
            nulSEL = "SELECTED"
            else
            nulSEL = ""
            end if
            %>
            <option value=0 <%=nulSEL %>>Alle</option>
            
            <%
            end if
            
            oRec.open strSQLaft, oConn, 3
				while not oRec.EOF
				
				if cint(aftid) = cint(oRec("sid")) then
				isSelected = "SELECTED"
				vlgtAft = oRec("navn") & " ("& oRec("aftalenr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("sid")%>" <%=isSelected%>><%=oRec("navn")%> (<%=oRec("aftalenr")%>) - <%=oRec("antalfak") %></option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
                </select>
                <%else %>
                <%=vlgtAft %>
                <%end if %>
        
        </td></tr>
	    
	    </tr>
        <tr><td valign=top><b>Job:</b>
        <br />Kun kontakter, job og aftaler<br /> med fakturaer på vises.</td><td valign=top>
	    <%
	    if cdbl(aftid) > 0 then
	    aftidKri = " AND serviceaft = " & aftid
	    else
	    aftidKri = ""
	    end if
	    
	    if kundeid <> 0 then
	    strjobKidSQL = "jobknr = " & kundeid
	    else
	    strjobKidSQL = "jobknr <> 0 "
	    end if
	    
	    strSQLjob = "SELECT j.id AS jid, j.jobnavn, j.jobnr, COUNT(f.fid) AS antalfak "_
	    &" FROM job j "_
	    &" LEFT JOIN fakturaer f ON (f.jobid = j.id AND shadowcopy = 0) "_
	    &" WHERE "& strjobKidSQL &" "& aftidKri &" AND f.fid <> 0 GROUP BY j.id ORDER BY j.jobnavn"
	  
	  'Response.Write strSQLjob
	  'Response.flush
	  
	        if print <> "j" then %>
            <select id="FM_job" name="FM_job" size=8 style="font-size : 9px; width:325px;">
             <%if cint(jobid) = -1 then
            ingSEL = "SELECTED"
            else
            ingSEL = ""
            end if
            %>
            
            <option value="-1" <%=ingSEL %>>Ingen</option>
            
            <%if cint(jobid) = 0 then
            nulSEL = "SELECTED"
            else
            nulSEL = ""
            end if
            %>
            
            <option value=0 <%=nulSEL %>>Alle</option>
            <%
            end if
            
            oRec.open strSQLjob, oConn, 3
				while not oRec.EOF
				
				if cint(jobid) = cint(oRec("jid")) then
				isSelected = "SELECTED"
				vlgtJob = oRec("jobnavn") & " ("& oRec("jobnr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("jid")%>" <%=isSelected%>><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) - <%=oRec("antalfak") %></option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				%>
         
        <% if print <> "j" then%>   
        </select><br /><input id="FM_jobikkedelafaftale" name="FM_jobikkedelafaftale" value="1" type="checkbox" <%=jidaftCHK%> />Vis kun job der ikke er del af aftale.
        
        <%else%>
        <%=vlgtJob %>
        <%end if %>
        </td></tr>
       
      
	  <tr>
      <tr>
       <td>&nbsp;</td>
		<td valign=top>
        <br />
        <b>Vælg måned:</b> 
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
	&nbsp;
        
     <b>År:</b>&nbsp;
	<select name="seomsfor" id="seomsfor" style="width:85px;">
			
			<%
			useaar = 2001
			for x = 0 to 20
			useaar = useaar + 1
			
			if cint(strYear) = cint(useaar) then
			aChk = "SELECTED"
			else
			aChk = ""
			end if
			%>
			<option value="<%=useaar%>" <%=aChk%>><%=useaar%></option>
			<%
			'x = x + 1
			next%>
		</select>

	</td>
	</tr>
    <tr>
	  <td valign=top>
	    <br /><b>Søg:</b>&nbsp;<br />
	    Faktura nr., aft.nr eller jobnr.<br />
	    Ignorerer periode og valgt kontakt.</td><td>
	    <br /> 
	    <%if print <> "j" then %>
	    <input id="FM_sog" name="FM_sog" value="<%=showSog%>" type="text" style="font-size : 9px; width:285px;" /><br />
	    Faktura nr, eks. 54 - 103, % = wildcard <br />
	    <%else %>
	    <%=showSog %>
	    <%end if %>
	    </td>
	    <td align=right>
	    <%if print <> "j" then %>
	    <input id="Submit1" type="submit" value="Vis afstemning >> " />
	    <%end if %></td>
	 </tr>
	 
	 
	 </form></table>
	 </div>
	 
	 </td></tr>
	 </table>
	 </div>

















		
				
				
				<%if print = "j" then%>
			
		
				
				
				
				
			
				
				<%
				strMrd = Request.Cookies("datoer")("st_md")
				strDag = Request.Cookies("datoer")("st_dag")
				strAar = Request.Cookies("datoer")("st_aar") 
				strDag_slut = Request.Cookies("datoer")("sl_dag")
				strMrd_slut = Request.Cookies("datoer")("sl_md")
				strAar_slut = Request.Cookies("datoer")("sl_aar")
				end if%>
		
		





<%if print = "j" then%>
<b>Periode:</b>
<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>
<br>
<br>
<%end if





sub faktr
saldo = saldo + oRec3("enheder")
%>

Viser kun fakturerbarer timer på de valgte job og aftaler.

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
		'Response.flush
        	
		t = 0	
		saldo = 0
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		
			'call tdbgcol_to_1(t)
			%>
			<br />
			<div style="position:relative; background-color:#FFFFFF; padding:10px 10px 10px 10px; border:1px silver solid;">
			<table cellpadding=0 cellspacing=0 border=0 width=600><tr>
			
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
			<tr><td><b>Job tilknyttet:</b></td><td>
			
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
			<br /><br />
			<table cellspacing=1 cellpadding=2 border=0 bgcolor="#cccccc" width=95%>
			
            <tr>
            <td colspan=5 bgcolor="#FFFFFF" style="height:20px;"><b>Stamdata</b></td>
            <td colspan=3 bgcolor="#FFFFE1" style="height:20px;"><b>Rest Estimat</b></td>
            <td colspan=8 bgcolor="#8caae6"><b>Realiseret Oms.</b> (kun fakturerbare timer)</td>
            <td colspan=11 bgcolor="yellowgreen"><b>Faktureret</b></td>
            </tr>
			
			<tr>
				
                <td valign=bottom bgcolor="#FFFFFF" style="padding:0px 2px 0px 5px;" class=lille><b>Job / Aftale</b></td>
                <td valign=bottom bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;" class=lille><b>Kontakt</b></td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Forecast</b></td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Forkalk. tim</b></td>

                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Brutto Oms.</b></td>
                
                 <td valign=bottom bgcolor="#FFFFE1" style="padding:0px 2px 0px 2px;" class=lille><b>Rest Estimat (Stade %)</b><br />Job er % afsluttet</td>
                <td valign=bottom bgcolor="#FFFFE1" style="padding:0px 2px 0px 2px;" class=lille><b>Forv. samlet tidsforbrug</b></td>
                <td valign=bottom bgcolor="#FFFFE1" style="padding:0px 2px 0px 2px;" class=lille><b>Balance</b><br />
                (Afv. %)</td>
                

                <td valign=bottom bgcolor="#EFf3FF" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Realiseret Timer</b><br />Periode</td>
                <td valign=bottom bgcolor="#8caae6" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Realiseret Timer</b><br />Akkumuleret</td>
              

                
                <!-- 
                fordelt på type
                Lnb. timer
                Fastpris
                Aftaler klippekort
                Aftaler Periode
                 -->
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Lbn. timer</td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Fastpris</td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Aftaler klippekort</td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Aftaler periode</td>

                <td valign=bottom bgcolor="#EFf3FF" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Realiseret Oms.</b><br />Periode</td>
                <td valign=bottom bgcolor="#8caae6" align=right style="padding:0px 2px 0px 2px;" class=lille><b>Realiseret Oms.</b><br />Akkumuleret</td>

                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Lbn. timer</td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Fastpris</td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Aftaler klippekort</td>
                <td valign=bottom bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;" class=lille>Aftaler periode</td>

                 <td valign=bottom bgcolor="#DCF5BD" style="padding:0px 2px 0px 2px;" class=lille><b>Faktureret timer</b><br />Periode</td>
                <td valign=bottom bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;" class=lille><b>Faktureret beløb</b><br />Periode</td>

                <td valign=bottom bgcolor="#DCF5BD" style="padding:0px 2px 0px 2px;" class=lille><b>Faktureret timer</b><br />Akkumuleret</td>
                <td valign=bottom bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;" class=lille><b>Faktureret beløb</b><br />Akkumuleret</td>
                <td valign=bottom bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;" class=lille><b>Balance</b><br />Forkalk. Timer/Faktureret timer ialt</td>
                <td valign=bottom bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;" class=lille><b>Balance</b><br />Brutto Oms./Faktureret Beløb ialt</td>
               <td valign=bottom bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;" class=lille><b>Betalingsplan</b></td>
			</tr>
			
			
			<%
			lastyear = 0
			
            '*** Hent fakturaer **'

            '*** Hent job der er faktureret + alle der ikke er faktureret, men der 
            '*** A) Ligger timer på i periode
            '*** B) Er fakmilepæle på
            '*** Hent aftaler der er faktureret + alle aftaler der
            '*** C) ligger fakmilepæle på
            
			'*** Fordel på typer lbn. timer, fastpris, aftaler klippekort, aftaler periode (partner aft.)
            '*** Vis realiserede timer alle job, men hvis kun timer til fakturering (forventet fakturering) hvis
            '*** D) Lbn timer (faktureres løbende når opgave er afsluttet = slutdato)
            '*** E) Fastpris og der findes en betalingsplan (fak. milepæle) min. 50% start og 50% ved afslutning - Oprettes auto ved joboprettelse
            '*** F) Job er en del af en aftale (= faktureres som en aftale / eller som job) og der findes en betalingsplan (fak. milepæle, se E)
            '*** G) Hvis fak. milpæle er slettet bliver forventet fak. = jobslutdato.
            '*** H) Aftaler kan ikke oprettes uden en betalingsplan

            strSQLmain = "SELECT j.jobnavn, j.jobnr, k.kkundenavn, k.kkundnr FROM job WHERE jobstartdato BETWEEN "

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
					<td colspan=27 style="padding-left:5px;"><b> <%=monthname(usemd)%> <%=useyear%></b></td>
			    </tr>
			<%
			end if
			
			lastyear = useyear
			
			
						'*** Aftale 
						'saldo = saldo - oRec("enheder")
						call tdbgcol_to_1(x)%>
						<tr>
							<td class=lille bgcolor="#FFFFFF" style="padding:0px 2px 0px 5px; width:200px;"><b>DSB Analyse</b> (12229)<br />
                            Lbn. timer</td>
							
							<td class=lille bgcolor="#FFFFFF" style="padding:0px 2px 0px 5px;">DSB first (1233333)</td>
							
                            <td align=right class=lille bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;">105,00</td>
							<td align=right class=lille bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;">75,00</td>
							<td align=right class=lille bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px; white-space:nowrap;">20.000,00 <%=basisValISO %></td>
							
							
                            <td class=lille bgcolor="#FFFFE1" style="padding:0px 2px 0px 5px;" align=right>20%</td>
                            <td class=lille bgcolor="#FFFFE1" style="padding:0px 2px 0px 5px;" align=right>200</td>
							<td class=lille bgcolor="#FFFFE1" style="padding:0px 2px 0px 5px;" align=right>35 t. / 21%</td>
                            
							
							
							
							
							<td class=lille bgcolor="#EFF3FF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(overfortsaldo, 2) %>
							<%overfortsaldo = 0 %>
							</td>
							
							<td class=lille bgcolor="#8caae6" align=right style="padding:0px 2px 0px 2px;">
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
							
							
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">
							<%
							saldomd = (useEnhTilSaldo - enhederRealiseret)
							saldoAkk = (saldoAkk +(saldomd)) 
							Response.write formatnumber(saldomd, 2)
							%>
							</td>
							
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">
							<b><%=formatnumber(saldoAkk, 2)%></b>
							</td>
							

                            <td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;">
							<b><%=formatnumber(saldoAkk, 2)%></b>
							</td>
							
							
							<td class=lille align=right bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;">
								
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

                            <td class=lille align=right bgcolor="#EFf3FF" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_Fak = (FakenhederPrMd - useEnhTilSaldo)
							Response.write formatnumber(enhSaldoTild_Fak, 2) 
							%>
							</td>
							
							<td class=lille align=right bgcolor="#8caae6" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_Fak = (FakenhederPrMd - useEnhTilSaldo)
							Response.write formatnumber(enhSaldoTild_Fak, 2) 
							%>
							</td>
							<td class=lille align=right bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoTild_FakAkk = (enhSaldoTild_FakAkk + (enhSaldoTild_Fak))
							Response.write formatnumber(enhSaldoTild_FakAkk, 2) 
							%>
							</td>
							
							<td class=lille align=right bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;">
							<%
							enhSaldoReal_Fak = (enhederRealiseret - FakenhederPrMd)
							enhSaldoReal_FakAkk = (enhSaldoReal_FakAkk + (enhSaldoReal_Fak))
							  
							%>
							<%=formatnumber(enhSaldoReal_FakAkk, 2)%>
							</td>
							
							
						
                            <td class=lille align=right bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;"><b><%=formatcurrency(FakKrprMd, 2)%></b></td>
							<td class=lille align=right bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px;">
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
							
							
							<td class=lille align=right bgcolor="#DCF5BD" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_Fak = (FakKrprMd - usePrisTilSaldoprMd)
							Response.write formatnumber(krSaldopris_Fak, 2) 
							%>
							</td>
							<td class=lille align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_FakAkk = (krSaldopris_FakAkk + (krSaldopris_Fak))
							%>
							<b><%=formatcurrency(krSaldopris_FakAkk, 2) %></b>
							</td>

                            <td class=lille align=right bgcolor="#DCF5BD" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_Fak = (FakKrprMd - usePrisTilSaldoprMd)
							Response.write formatnumber(krSaldopris_Fak, 2) 
							%>
							</td>
							<td class=lille align=right bgcolor="yellowgreen" style="padding:0px 2px 0px 2px;">
							<%
							krSaldopris_FakAkk = (krSaldopris_FakAkk + (krSaldopris_Fak))
							%>
							<b><%=formatcurrency(krSaldopris_FakAkk, 2) %></b>
							</td>
							
                            <td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(timerThis, 2)%></td>

                            <td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(timerThis, 2)&" "&basisValISO %></td>
							<!--
							<td class=lille bgcolor="#FFFFFF" align=right style="padding:0px 2px 0px 2px;"><%=formatnumber(timerThis, 2)%></td>
                            -->
							
                            <td class=lille bgcolor="#FFFFFF" style="padding:0px 2px 0px 2px; white-space:nowrap;">21. sep. 2010<br />21. okt 2010</td>
                            
                            <!--
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
							
							-->
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

<br /><br />
            &nbsp;



		</div>
	<%end if%>
	<br /><br />
            &nbsp;
<!--#include file="../inc/regular/footer_inc.asp"-->