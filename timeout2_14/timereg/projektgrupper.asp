<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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
	<!--include file="../inc/regular/topmenu_inc.asp"-->
	<% 
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en projektgruppe.<br>"_
    &"Du vil samtidig slette alle tilhørsforhold til denne gruppe.<br>"_
	&"Er du sikker?<br>"
    slturl = "projektgrupper.asp?menu=job&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM projektgrupper WHERE id = "& id &"")
	Response.redirect "projektgrupper.asp?menu=job"
	
	case "opdprorel"
	
	z = 0
	x = Request("FM_antal_x")
	y = Request("FM_antal_y")
	
	'Response.write "x " & x & " y " & y & "<br>"
	strProjektgruppeId = request("FM_projektgruppeId")
	
	'** Finder jobid på job hvor denne projktgruppe har adgang til ***'
    guidjobids = ""
    guideasyids = ""
	call projktgrpPaaJobids(strProjektgruppeId)
	
	call positiv_aktivering_akt_fn() 'wwf
                         
	
		            
	'oConn.execute("DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &"")
	
	
	For z = 0 to y - 1
	'Response.write z & "<br>"
	'Response.write "strMid: " & request("FM_Mid_"&z&"")& " , Optages: "& request("FM_medlem_"&z&"") & "<br>"
	'Response.flush
		
		strMid = request("FM_Mid_"&z&"")
		    
		    
		   if len(trim(strMid)) <> 0 then 
		    
		    fandtes = 0
		    strSQLfindes = "SELECT projektgruppeId, MedarbejderId"_
		    &" FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId & " AND MedarbejderId = "&strMid
		    
		    oRec4.open strSQLfindes, oConn, 3
		    if not oRec4.EOF then
		    
		    fandtes = 1
		    
		    end if
		    oRec4.close  
		    
		    'Response.Write "fandtes:" & fandtes
		    
		    '** Team leder **'
		    if request("FM_teamleder_"&z&"") = 1 then
    		teamleder = 1
    		else
    		teamleder = 0
    		end if
		    
            '** Notificer **'
		    if request("FM_notificer_"&z&"") = 1 then
    		notificer = 1
    		else
    		notificer = 0
    		end if


		    '** fandtes relation i forvejen ??***'
		    if cint(fandtes) = 1 then
		    
		        '*** Hvis ja, skal den slettes? ***'
		        if request("FM_medlem_"&z&"") <> "1" then
    		      oConn.execute("DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &" AND MedarbejderId = "& strMid &"")
                  'Response.Write " -- DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &"" & "<br

                  '*** Sletter fra timereg usejob og tilføjer igen **'
                  del = 1 '1 slet fra timereg_usejob / 0: tilføj kun
                  call tilfojtilTU(strMid, del)
                
                else
                 
                 oConn.execute("UPDATE progrupperelationer SET teamleder = "& teamleder &", notificer = "& notificer &" WHERE projektgruppeId = "& strProjektgruppeId &" AND MedarbejderId = "& strMid &"")  

                 end if
		    
		    else
		    
		         '*** skal den oprettes? ***
		         if request("FM_medlem_"&z&"") = 1 then
    		        
    		        
    		        
		            oConn.execute("INSERT INTO progrupperelationer (projektgruppeId, MedarbejderId, teamleder, notificer) VALUES ("& strProjektgruppeId &", "& strMid &", "& teamleder &", "& notificer &")")
		            
		            
		            '*** sætter aktiv i guiden aktiv job
                    'forvalgt = 0 'off / 1 = on
                    'Response.Write guideasyids
                    'Response.end
		            'call setGuidenUsejob(strMid, guidjobids, 1, guideasyids, forvalgt,2)
                    del = 0 '1 slet fra timereg_usejob / 0: tilføj kun
                    
                    if cint(positiv_aktivering_akt_val) <> 1 then
                    call tilfojtilTU(strMid, del)
		            end if
		            
		            
		         end if
		    
		    
		    
		    end if
		
		end if 'strmid <> 0
	
	Next
	
	               
	'Response.end
	
	Response.redirect "projektgrupper.asp?menu=job&lastid="&strProjektgruppeId
	
	case "med"
	%>
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

    <%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">

	<%
	tTop = 0
	tLeft = 0
	tWdth = 800
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	
	<%
	strSQL = "SELECT id, navn FROM projektgrupper WHERE id=" & id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	gruppeNavn = oRec("navn")
	end if
	oRec.close
	%>

        <h3>Projektgrupper</h3>
	
	
	Medarbejdere i projektgruppen <b><%=gruppeNavn%></b>:
     <form action="projektgrupper.asp?menu=job&func=opdprorel" method="post">
	
	
	<input type="hidden" name="FM_projektgruppeId" value="<%=id%>">
	
	
	<%
    call progrpHeader

	Dim MedarbId()
	'*** Henter alle medlemmer
	strSQL = "SELECT Mnavn, Mid, MedarbejderId, teamleder, init, mnr, mansat, notificer FROM progrupperelationer, medarbejdere"_
    &" WHERE Mid = MedarbejderId AND ProjektgruppeId = "&id&" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	
	x = 0
	Redim MedarbId(x)
	while not oRec.EOF 
	
	select case right(x, 1)
	case 0,2,4,6,8
	bgt = "#ffffff"
	case else
	bgt = "#D6Dff5"
	end select
	
	
	Redim preserve MedarbId(x)
	MedarbId(x) = oRec("Mid")%>
	<tr>
		<td bgcolor="#cccccc" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgt %>">
		<td height="20">
        
         <%if level <= 2 then %>
         <a href="medarb_red.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>" class=vmenu target="_blank"><%=oRec("Mnavn")%> (<%=oRec("Mnr")%>)
            <% 
            if len(trim(oRec("init"))) <> 0 then
            %>
            - <%=oRec("init") %>
            <%end if    
            %>
         </a>
         <%else %>
         <%=oRec("Mnavn")%> (<%=oRec("Mnr")%>)
            <% 
            if len(trim(oRec("init"))) <> 0 then
            %>
            - <%=oRec("init") %>
            <%end if    
            %>

         <%end if %>


        <%if oRec("mansat") = "2" then %>
        - Deaktiveret
         <%end if %>

         <%if oRec("mansat") = "3" then %>
        - Passiv
         <%end if %>
        
        </td>
		<td><input type="checkbox" name="FM_medlem_<%=x%>" value="1" CHECKED>
		<input type="hidden" name="FM_Mid_<%=x%>" value="<%=oRec("Mid")%>"></td>
		
		<%if cint(oRec("teamleder")) =  1 then
		tmCHK = "CHECKED"
		else
		tmCHK = ""
		end if %>
		
        <%if cint(oRec("notificer")) =  1 then
		ntCHK = "CHECKED"
		else
		ntCHK = ""
		end if %>

		<%if level = 1 then %>
		<td><input type="checkbox" name="FM_teamleder_<%=x%>" value="1" <%=tmCHK %>></td>
        <td><input type="checkbox" name="FM_notificer_<%=x%>" value="1" <%=ntCHK %>></td>
		<%else %>
		<td><input type="checkbox" name="" value="1" DISABLED <%=tmCHK %>>
        <input id="Hidden1" name="FM_teamleder_<%=x%>" value="<%=oRec("teamleder") %>" type="hidden" /></td>
        <td><input type="checkbox" name="" value="1" DISABLED <%=ntCHK %>>
        <input id="Hidden1" name="FM_notificer_<%=x%>" value="<%=oRec("notificer") %>" type="hidden" /></td>
		<%end if %>
    </tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close
	%>
	
	</table>
	
	<br /><br />
	
	<b>Medarbejdere uden for projektgruppe:</b> (aktive og passive)
	
	
	
	
	<%

    call progrpHeader

	'*** Henter ikke medlemmer
	Dim T
	T = 0
	For T = 0 to x - 1
	strExclude = strExclude & "Mid <> "&MedarbId(T)&" AND "
	Next
	
	antChar = len(strExclude)
	if antChar > 0 then
	LantChar = left(strExclude, (antChar -5)) 
	strExcludeFinal = "WHERE " & LantChar
	
	strSQL = "SELECT Mnavn, Mid, init, mnr, mansat FROM medarbejdere "& strExcludeFinal &" AND (mansat <> '2') ORDER BY Mnavn"
	else
	strSQL = "SELECT Mnavn, Mid, init, mnr, mansat FROM medarbejdere WHERE mansat <> '2' ORDER BY Mnavn"
	end if
	
	oRec.open strSQL, oConn, 3
	
	y = x
	r = oRec.recordcount
	
	while not oRec.EOF 
	
	select case right(y, 1)
	case 0,2,4,6,8
	bgt = "#ffffff"
	case else
	bgt = "#D6Dff5"
	end select
	
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgt %>">
		<td height="20">
         
         <%if level <= 2 then %>
         <a href="medarb_red.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>" class=vmenu target="_blank"><%=oRec("Mnavn")%> (<%=oRec("Mnr")%>)
            <% 
            if len(trim(oRec("init"))) <> 0 then
            %>
            - <%=oRec("init") %>
            <%end if    
            %>
         </a>
         <%else %>
         <%=oRec("Mnavn")%> (<%=oRec("Mnr")%>)
            <% 
            if len(trim(oRec("init"))) <> 0 then
            %>
            - <%=oRec("init") %>
            <%end if    
            %>

         <%end if %>

         <%select case oRec("mansat") 
         case "2"%>
         - Deaktiveret
         <%case "3"  %>
        - Passiv
         <%end select %>
         
         </td>
		<td><input type="checkbox" name="FM_medlem_<%=y%>" value="1">
		<input type="hidden" name="FM_Mid_<%=y%>" value="<%=oRec("Mid")%>"></td>
		
        <%if level = 1 then %>
		<td><input type="checkbox" name="FM_teamleder_<%=y%>" value="1"></td>
        <td><input type="checkbox" name="FM_notificer_<%=y%>" value="1"></td>
		<%else %>
		<td><input type="checkbox" name="" value="0" DISABLED>
        <input id="Text1" name="FM_teamleder_<%=y%>" value="0" type="hidden" /></td>

        <td><input type="checkbox" name="" value="0" DISABLED>
        <input id="Text1" name="FM_notificer_<%=y%>" value="0" type="hidden" /></td>
		<%end if %>

	</tr>
	<%
	y = y + 1
	oRec.movenext
	wend
	oRec.close%>
	
	
      


    </table>
    
    </div>
    <br /><br />
 


    	<input type="hidden" name="FM_antal_y" value="<%=y%>">
	<img src="ill/blank.gif" width="500" height="1" border="0" alt="">
  
	<input type="submit" value="Opdater gruppe >>" />
	</form>		
	<br><br>

 
 
     <a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
    <br>
    <br>

	

	</div>
	<br /><br /><br /><br /><br />&nbsp;
	
	
	<%
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
	
		<%
		errortype = 8
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
		
		if len(trim(request("FM_opengp"))) <> 0 then
		intopengp = 1
		else
		intopengp = 0
		end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO projektgrupper (navn, editor, dato, opengp) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intopengp &")")
		
			strSQL = "SELECT id FROM projektgrupper WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			
			lastid = oRec("id")
			
			end if
			
			oRec.close
		
		else
		
		oConn.execute("UPDATE projektgrupper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', opengp = "& intopengp &" WHERE id = "&id&"")
		lastid = id
		
		end if


        if func = "dbopr" AND request("FM_tilfoj_m") = "1" then
        '*** tilføjer den der opretter gruppen som teamleder ***
        strSQLt = "INSERT INTO progrupperelationer (projektgruppeid, medarbejderid, teamleder) VALUES ("& lastId &", "& session("mid") &", 1)"
        oConn.execute(strSQLt)
        end if

		
		Response.redirect "projektgrupper.asp?menu=job&lastid="&lastid
		end if
	
	
        
        
    
    
    
    case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato, opengp FROM projektgrupper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intOpenGp = oRec("opengp")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

    <%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->

	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">

        <%
	tTop = 0
	tLeft = 0
	tWdth = 800
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>

	<h3>Projektgrupper <span style="font-size:9px; font-weight:lighter;"><%=varbroedkrumme%></span></h3>
	<table cellspacing="0" cellpadding="5" border="0" width="100%" style="background-color:#FFFFFF;"">
	<form action="projektgrupper.asp?menu=job&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">

	<tr bgcolor="#5582d2" >
		<td colspan=2 valign="top" class="alt">
		<%if dbfunc = "dbred" then%>
		
		Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b>
		<%end if%></td>
	</tr>
	
	<tr>
	
		
		<td colspan="2"><br />Navn:&nbsp;&nbsp;<input type="text" name="FM_navn" value="<%=strNavn%>" style="width:200px;"></td>
		
	<%if intOpenGp = "1" then
	intOpenGpCHK = "CHECKED"
	else
	intOpenGpCHK = ""
	end if %>
	
	<tr>
	
		<td colspan="2" style="padding-left:30px;">
            <input id="Checkbox1" name="FM_opengp" value="1" type="checkbox" <%=intOpenGpCHK %>/>Åben gruppe (medarbejdere må selv tilføje sig)</td>
	
	</tr>

    <%if func = "opret" then %>
    <tr>
	
		<td colspan="2" style="padding-left:30px;">
            <input id="Checkbox2" name="FM_tilfoj_m" value="1" type="checkbox"  CHECKED/>Tilføj dig selv som teamleder i gruppen</td>
		
	</tr>

    <%end if %>
	<tr>
		
		<td colspan="2" align="right"><br /><br /><input type="submit" value="Opdater >>"></td>
		
	</tr>
	
	</form>
	</table>
    </div><!-- tablediv -->
	<br><br>
	<br>
	<a href="Javascript:history.back()"><< Tilbage</a>
	<br>
	<br>
	</div>
	<%case else
	
	
	if len(request("lastid")) <> 0 then
	lastid = request("lastid")
	else
	lastid = 0
	end if %>
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	


    <%call menu_2014() %>

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px; visibility:visible;">
	
	<%
	oimg = "blank.gif"
	oleft = 0
	otop =  0
	owdt = 400
	oskrift = "Projektgrupper"
	
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	'call sideoverskrift_2014(oleft, otop, owdt, oskrift)
	
	
    
	
	tTop = 0
	tLeft = 0
	tWdth = 900
	
	
	call tableDiv(tTop,tLeft,tWdth)


        %><table width="100%"><tr><td width="82%">
            <%
        call sideoverskrift_2014(oleft, otop, owdt, oskrift)


          if level <= 2 OR level = 6 then

            %>

                     </td><td><%
      
                nWdt = 180
                nTxt = "Opret nyt projektgruppe"
                nLnk = "projektgrupper.asp?menu=job&func=opret&id=0"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt)
        end if    
	
	%>
            </td></tr></table>
        <br><br>
	<table cellspacing="0" cellpadding="6" border="0" width="100%" bgcolor="#FFFFFF">
	<tr bgcolor="#5582D2">
		<td height=30 class="alt">Id - <b>Projektgruppe</b><br />(Antal medl. aktive + passive)</td>
		<td class="alt"><b>Tilføj/fjern medlemmer</b></td>
		<td class="alt">Privat / Åben?</td>
		<td class="alt"><b>Slet proj.gruppe?</b></td>
	</tr>
	<%
	
	strSQL = "SELECT id, navn, opengp FROM projektgrupper WHERE id > 1 ORDER BY navn"
	oRec.open strSQL, oConn, 3
	
	x = 0
	n = 0
	while not oRec.EOF 
				
				x = 0
				call antalMediPgrp(oRec("id"))
				
				if antalMediPgrpX > 0 then
				x = antalMediPgrpX
				Antal = x
				else
				x = 0
				Antal = x
				end if
	
	            antalMediPgrpX = 0

                call fTeamleder(session("mid"), oRec("id"))
	            
	if cint(lastid) = oRec("id") then
	bgcol = "#ffff99"
	else
	    
	    select case right(n, 1)
	    case 0,2,4,6,8
	    bgcol = "#EFF3FF"
	    case else
	    bgcol = "#ffffff"
	    end select
	
	
	end if%>
	
	
	<tr bgcolor="<%=bgcol%>">
	<%if oRec("id") <> "10" AND (oRec("opengp") = 1 OR level = 1 OR erTeamleder = 1) then%>
	<td><%=oRec("id") %> - <a href="projektgrupper.asp?menu=job&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>&nbsp;&nbsp;</td>
	<td><a href="projektgrupper.asp?menu=job&func=med&id=<%=oRec("id")%>" class=vmenu>&nbsp;Se medlemmer </a> (<%=antal%>)</td>
	<%else
	Response.write "<td height=20>"& oRec("id") &" - "& oRec("navn") &"&nbsp;&nbsp;("&antal&")</td>"
	Response.write "<td>&nbsp;</td>"
	end if%>
	<td>
	<%if oRec("opengp") = 0 then %>
	Privat
	<%else %>
	Åben
	<%end if %>
	</td>
	
	<%if oRec("id") <> "10" AND (level = 1 OR erTeamleder = 1) then%>
	<td><a href="projektgrupper.asp?menu=job&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="" border="0"></a></td>
	<%else%>
	<td>&nbsp;</td>
	<%end if%>
	</tr>
	<%
	x = 0
	n = n + 1
	oRec.movenext
	wend
	%>
	
	</table>
	</div>
	
	
	<br><br>
	<br>
	<a href="Javascript:history.back()"><< Tilbage</a>
	<br>
	<br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
