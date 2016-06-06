<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
   
    
    

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210px; top:150px; visibility:visible; background-color:#FFFFFF; border:10px #CCCCCC solid; padding:20px;"">
	<h3>Materialegruppe - slet?</h3>
	
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en materiale-gruppe. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="materiale_grp.asp?menu=mat&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM materiale_grp WHERE id = "& id &"")
	Response.redirect "materiale_grp.asp?menu=mat&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(trim(request("FM_navn"))) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = 104
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		
		strNavn = SQLBless(request("FM_navn"))
		
		
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		strGrpNr = SQLBless(request("FM_nr"))
        medarbansv = request("FM_medarbansv")

        matgrp_konto = request("FM_matgrp_konto")
		
		
		 %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(request("FM_av")) 
			if isInt > 0 OR len(trim(request("FM_av"))) = 0 then
				
				%>
				<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
				<%
				
				errortype = 105
				call showError(errortype)
			
			isInt = 0
			
			Response.end
			end if
			
		
		strAv = replace(formatnumber(request("FM_av"), 0), ",",".")
        kundeid = request("FM_kunde")
		
		if func = "dbopr" then
		
		oConn.execute("INSERT INTO materiale_grp (navn, editor, dato, nummer, av, mgkundeid, medarbansv, matgrp_konto) VALUES "_
		&" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strGrpNr &"', "& strAv &", "& kundeid &", "& medarbansv &", "& matgrp_konto &")")
		
		else
		
		oConn.execute("UPDATE materiale_grp SET navn ='"& strNavn &"', "_
		&" nummer = '"& strGrpNr &"', editor = '" &strEditor &"', dato = '" & strDato &"', av = "& strAv &", mgkundeid = "& kundeid &", medarbansv = "& medarbansv &", matgrp_konto = "& matgrp_konto &" WHERE id = "&id&"")
		
		    
		    if request("FM_opdater_av") = "1" then
	        strSQLupd = "UPDATE materialer SET salgspris = (indkobspris * (("& strAv &"/100) + 1)) WHERE matgrp = " & id
	        
	        oConn.execute(strSQLupd)
	        end if
		    
		
		end if
		
		
		
		
		Response.redirect "materiale_grp.asp?menu=mat&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	strGrpNr = ""
	sidst_redigeret = ""
	strAv = 0
    medarbansv = 0
	matgrp_konto = 0

	else
	strSQL = "SELECT navn, editor, dato, nummer, av, mgkundeid, medarbansv, matgrp_konto FROM materiale_grp WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strGrpNr = oRec("nummer")
	strAv = oRec("av")
    kid = oRec("mgkundeid")
    medarbansv = oRec("medarbansv")
    matgrp_konto = oRec("matgrp_konto")
	
	end if
	oRec.close
	
	sidst_redigeret = "Sidst opdateret den <b>"&formatdatetime(strDato, 1)&"</b> af <b>"&strEditor&"</b>"
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
        -->



    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible; background-color:#ffffff; padding:20px;">
	<h4>Materiale / produktgrupper <span style="font-size:9px;">- <%=varbroedkrumme %></span></h4>
	<table cellspacing="0" cellpadding="2" border="0" width="500" bgcolor="#ffffff">
	<form action="materiale_grp.asp?menu=mat&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	
  	<tr>
		
		<td colspan="2" valign="middle" style="height:30px;"><%=sidst_redigeret%>&nbsp;</td>
		
	</tr>

        	<tr>
		
		<td colspan="2" valign="middle" style="height:30px;">&nbsp;</td>
		
	</tr>
	<tr>
		
		<td><font color=darkred>*</font> Navn:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" ></td>
		
	</tr>
	<tr>
		
		<td>Gruppe nummer / betegnelse:</td>
		<td><input type="text" name="FM_nr" value="<%=strGrpNr%>" size="30" ></td>
		
	</tr>
    <tr>
		
		<td>Kunde:</td>
		<td> 
      <select name="FM_kunde" id="FM_kunde" size="1" style="width:205px;">
          <option value="0">Ingen</option>
		<%
		
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				
		</select></td>
		
	</tr>
	<tr>
		
		<td>Avance %:</td>
		<td><input type="text" name="FM_av" value="<%=strAv%>" size="10" > % (heltal)</td>
		
	</tr>
	<tr>
		
		<td colspan=2 style="height:40px;">&nbsp;<input id="FM_opdater_av" name="FM_opdater_av" type="checkbox" value="1" /> Opdater alle materialer i denne gruppe til at følge den valgte avance.<br /><br />&nbsp;</td>
		
	</tr>
    <tr><td>Medarbejder ansvarlig: <br /><span style="color:#999999;">Bliver notificeret når <br />minimumslager <br />bliver overskreddet</span>
        </td><td valign="top">
        <%
       strSQL = "SELECT Mid, Mnavn, init FROM medarbejdere WHERE mansat = 1 GROUP BY mid ORDER BY Mnavn"
					
                 

					%>
					<select name="FM_medarbansv" style="width:205px;">
                    <option value="0">Vælg medarbejder..</option><!-- onchange="submit(); -->
					<%
					
					oRec.open strSQL, oConn, 3
					while not oRec.EOF 
					
					if cint(oRec("Mid")) = cint(medarbansv) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if
                    
                        if len(trim(oRec("init"))) <> 0 then
                        medTxt = oRec("mnavn") & " ["& oRec("init") &"]"
                        else
                        medTxt = oRec("mnavn") 
                        end if
                        
                    %>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=medTxt%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>

        </td></tr>
	
	<tr>
      <tr><td>Konto:</td><td valign="top">
        <%
        call licKid()

        '*** Kontoplan
        if licensindehaverKid <> 0 then
        licensindehaverKid = licensindehaverKid
        else
        licensindehaverKid = 0
        end if

        
         strSQLKontoliste = "SELECT navn, kontonr, id FROM kontoplan "_
         &" WHERE kid = "& licensindehaverKid &" ORDER BY navn"
         
					
                 

					%>
					<select name="FM_matgrp_konto" style="width:205px;">
                    <option value="0">Vælg konto..</option><!-- onchange="submit(); -->
					<%
					
					oRec.open strSQLkontoliste, oConn, 3
					while not oRec.EOF 
					
					if cint(matgrp_konto) = cint(oRec("id")) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if
                    
                        
                    %>
					<option value="<%=oRec("id")%>" <%=rchk%>><%=oRec("navn") & " ("&oRec("kontonr")%>)</option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>

        </td></tr>


		<td colspan="4" align=right><br><br><input type="submit" value="<%=varSubVal%> >>"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><< Tilbage</a>
	<br>
	<br>
	</div>
	
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
    -->

    <%call menu_2014() %>

	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; width:1000px; visibility:visible; padding:20px; background-color:#ffffff;">
	<h4>Materiale / produktgrupper</h4>
	<table cellspacing="0" cellpadding="0" border="0" width="100%"><form action="materiale_grp.asp?menu=mat" method="post">
	<tr>
		<!--
		<td><b>Søg efter gruppe:</b>&nbsp;&nbsp;(Gruppenavn ell. nummer)<br>
		<input type="text" name="FM_sog" id="FM_sog" value="<%=request("FM_sog")%>" style="width:200px;"> <input type="image" src="../ill/pilstorxp.gif"></td>
		-->
		<td align=right><a href="materiale_grp.asp?menu=mat&func=opret">Opret ny gruppe <img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br></td>
	</tr></form>
	</table>
	<br>
	<table cellspacing="0" cellpadding="4" border="0" width="100%">
	
	
	<tr bgcolor="#5582D2">
        <td><a href="materiale_grp.asp?menu=mat&sort=navn" class=alt><u><b>Gruppe og avance %</b></u></a></td>
        <td ><a href="materiale_grp.asp?menu=mat&sort=nr" class=alt><u><b>Grp. Betegnelse / nr.</b></u></a></td>
        <td class=alt><b>Kunde</b></td>
        <td class=alt><b>Konto</b></td>
		<td class=alt><b>Medarbejder<br />ansvarlig</b></td>
		<td class=alt><b>Antal varenr</b></td>
		
        <td class=alt>Slet</td>
        
	</tr>
	<%
	sort = Request("sort")

        strSQL = "SELECT mg.id, mg.navn, mg.nummer, mg.av, mg.mgkundeid, k.kkundenavn, kkundenr, medarbansv, m.mnavn, init, matgrp_konto, kp.kontonr, kp.navn AS kontonavn FROM materiale_grp mg "_
        &" LEFT JOIN kunder AS k ON (k.kid = mg.mgkundeid) "_
        &" LEFT JOIN kontoplan AS kp ON (kp.id = mg.matgrp_konto) "_
        &" LEFT JOIN medarbejdere AS m ON (m.mid = mg.medarbansv) WHERE mg.id > 0"

	if sort = "navn" then
	strSQL = strSQL & " ORDER BY mg.navn"
	else
	strSQL = strSQL & " ORDER BY mg.nummer"
	end if
	
	'Response.Write strSQL
	
	x = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	select case right(x, 1)
	case 0, 2, 4, 6, 8
	bgthis = "#ffffff"
	case else
	bgthis = "#eff3ff"
	end select%>
	
	<tr bgcolor="<%=bgthis%>">
		<td style="border-bottom:1px #CCCCCC solid;"><a href="materiale_grp.asp?menu=mat&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%></a>&nbsp;&nbsp;<%=oRec("av") %>%</td>
       
		<td style="border-bottom:1px #CCCCCC solid;"><%=oRec("nummer")%></td>
        <td style="border-bottom:1px #CCCCCC solid;">
            <%if oRec("mgkundeid") <> 0 then %>
            <%=oRec("kkundenavn") & " ("& oRec("kkundenr") &")" %>
            <%else %>
            &nbsp;
            <%end if %>
        </td>

        <td style="border-bottom:1px #CCCCCC solid;">
            <%if oRec("matgrp_konto") <> 0 then %>
            <%=oRec("kontonavn") & " ("& oRec("kontonr") &")" %>
            <%else %>
            &nbsp;
            <%end if %>
        </td>

		 <td style="border-bottom:1px #CCCCCC solid;">

            <% if len(trim(oRec("init"))) <> 0 then
                        medTxt = oRec("mnavn") & " ["& oRec("init") &"]"
                        else
                        medTxt = oRec("mnavn") 
                        end if %>

            <%=medTxt %>&nbsp;

        </td>

		<td style="border-bottom:1px #CCCCCC solid;">
		<%
		antal = 0
		strSQL2 = "SELECT count(m.id) AS antal FROM materialer m "_
	    &" WHERE (m.matgrp = "& oRec("id") &") GROUP BY m.matgrp"
	    oRec2.open strSQL2, oConn, 3
        if not oRec2.EOF then
        antal = oRec2("antal")
        end if
        oRec2.close
	
		%>
		<%=antal%></td>
		<td align="center" style="border-bottom:1px #CCCCCC solid;">
		<%if cdbl(antal) = 0 then %>
		<a href="materiale_grp.asp?menu=mat&func=slet&id=<%=oRec("id")%>" class="red">X</a>
		<%else %>
		&nbsp;
		<%end if %></td>
		
	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	%>	
	
	</table>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
