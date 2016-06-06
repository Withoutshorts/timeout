<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->



<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	

    call smileyAfslutSettings()

	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("lastid")) <> 0 then
	lastID = request("lastid")
	else
	lastID = 0
	end if 
	
	level = session("rettigheder")
	
	'Response.Write "level" & level
	
	if len(request("medarbSel")) <> 0 then
	medarbSel = request("medarbSel")
	else
	medarbSel = session("mid")
	end if
	
	if level <> 1 then
	showalle = 0
	medarbSQLkri = " m.mid = " & medarbSel
	else
	showalle = 1
	medarbSQLkri = " m.mid <> 0 "
	end if
	
	'if len(request("hidemenu")) <> 0 then
	'hidemenu = request("hidemenu")
	'else
	'hidemenu = 0
	'end if
	
	rdir = request("rdir")
		
	thisfile = "stat_korsel"
	
	'*** Smiley ***''
    call ersmileyaktiv()
	
	media = request("media")
	'printdo = request("print")
	
	
	function ArTilDatoKm()
	%>
	  År -> Dato (<%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2) %>) :
      
       
        
        <% 
        '** Sum år til dato ***'
        strSQLsum = "SELECT SUM(timer) AS kmsum, tfaktim FROM timer "_
        &" WHERE tmnr = "& ltmnr &" AND "_
        &" tdato BETWEEN '"& strAar_slut & "/1/1" &"' AND '"& sqlDatoSlut &"'"_
        &" AND tfaktim = 5 GROUP BY tmnr, tfaktim"
        kmAarDato = 0
        
        'Response.Write strSQLsum
        'response.flush
        
        oRec3.open strSQLsum, oConn, 3
        if not oRec3.EOF then
        
        kmAarDato = oRec3("kmsum")
        
        end if
        oRec3.close
	    
	    Response.Write "<b>"& formatnumber(kmAarDato, 2) & "</b>"
	
	end function
	
	
	
	if media <> "print" then
	%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	
	<%'call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call stattopmenu()
	'end if
	%>
	</div>

    -->
	<%

    call menu_2014()

	sideDivTop = 102
	sideDivLeft = 90
	
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	sideDivTop = 20
	sideDivLeft = 20
	
	end if
	
	if len(request("FM_sumprmed")) <> 0 AND request("FM_sumprmed") <> 0 then
	showTot = 1
	sumChk = "CHECKED"
	else
	sumChk = ""
	showTot = 0
	end if
	%>
	
	<div id="sindhold" style="position:absolute; left:<%=sideDivLeft%>px; top:<%=sideDivTop%>px; visibility:visible;">
	
	<% 
	oimg = "ikon_stat_korsel_48.png"
	oleft = 0
	otop = 0
	owdt = 600
	oskrift = "Kørsels statistik"
	
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	%>
	
	<%call filterheader_2013(0,0,600,oskrift)%>
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="stat_korsel.asp?menu=stat&FM_usedatokri=1&hidemenu=<%=hidemenu%>" method="post">
	<tr>
			<td valign=top><b>Medarbejder:</b><br>
			<%
			
			
			
			strSQL = "SELECT mnavn, mid, mnr FROM medarbejdere m WHERE m.mid <> 0 AND mansat <> 2 AND mansat <> 3 AND "& medarbSQLkri &" ORDER BY mnavn"
			'Response.write strSQL
			'Response.flush
			oRec.open strSQL, oConn, 3 
			
			if media <> "print" then
			
			%><select name="medarbSel" id="medarbSel" style="width:200px;">
			
			<%if showalle = 1 then%>
			<option value="0">Alle</option>
			<%end if
			
			else
			medarbPrVal = "Alle"
			end if
			
			
			while not oRec.EOF 
			
			 if cint(medarbSel) = oRec("mid") then
			 mSEL = "SELECTED"
			 medarbPrVal = oRec("mnavn") &"&nbsp;("& oRec("mnr")&")"
			 else
			 mSEL = ""
			 end if 
			 
			 
			if media <> "print" then%>
			<option value="<%=oRec("mid")%>" <%=mSEL%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)</option>
			<%
			end if
			
			oRec.movenext
			wend
			oRec.close
			
			
			if media <> "print" then
			 %>
			</select><br><br>
			<%else %>
			<%=medarbPrVal%>
			<%end if %>
			
			<%if media <> "print" then %>
			<input type="checkbox" name="FM_sumprmed" id="FM_sumprmed" value="1" <%=sumChk%>> Vis sum fordelt pr. medarbejder og 60 dages regel.<br />
			24 md. forud for slut dato i valgt interval.&nbsp;&nbsp;</td>
			<td valign=top><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"-->
			</td>
			<td>
                <input id="Submit2" type="submit" value="Kør >> " /></td>
	</tr></form>
	    <%else 'print 
	    dontshowDD = "1"
	    %>
	    
	    <!--#include file="inc/weekselector_s.asp"-->
	    <%end if %>
	</table>
	
	
	<!-- filter header -->
	</td></tr></table>
	</div>
	
	<%
	

	
	
	'*** Alle timer, uanset fomr. ***
	sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
	sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	datoStart = strDag&"/"&strMrd&"/"&strAar
	datoSlut = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
	
	if medarbSel <> 0 then
    medarbSQLkri = " tmnr = " & medarbSel
	else
	medarbSQLkri = " tmnr <> 0 "
	end if
	
	
	tTop = 20
	tLeft = 0
	tWdth = 1004
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	if showTot = 0 then
	
	%>
	    
	 <br />
        <b>Periode: <%=formatdatetime(datoStart, 1)%> til <%=formatdatetime(datoSlut, 1)%>  </b>
    <br /><br />
	
	   <table cellpadding=2 cellspacing=0 border=0 width=100%>
	   <tr bgcolor="#5582d2">
	    <td valign=bottom class=alt style="border-bottom:1px #cccccc solid;"><b>Dag</b></td>
        <td valign=bottom class=alt style="border-bottom:1px #cccccc solid;"><b>Medarbejder</b></td>
         <td valign=bottom class=alt style="border-bottom:1px #cccccc solid;"><b>Kontakt</b>(kunde)</b><br />Job</td>
        <td valign=bottom class=alt style="border-bottom:1px #cccccc solid;"><b>Aktivitet</b></td>
        <td valign=bottom class=alt style="width:200px; border-bottom:1px #cccccc solid;"><b>Kommentar</b></td>
        <td valign=bottom valign=bottom class=alt style="border-bottom:1px #cccccc solid;"><b>Bopælstur</b><br />
        (60 dages regel)</td>
        <td valign=bottom class=alt style="width:200px; border-bottom:1px #cccccc solid;"><b>Destination</b><br />
        (60 dages regel)</td>
        <td valign=bottom class=alt align=right style="border-bottom:1px #cccccc solid;"><b>Km</b></td>
   
   <%'*** Joblog denne uge ***' 
   
   strEksportTxt = "Dato;Medarbejder;Kontakt;Job;Aktivitet;Bopælstur;Destination;Km;Kommentar;xx99123sy#z"
  
   strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
   &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, timerkom, bopal, "_
   &" m.mnr, m.init, m.mid, destination FROM timer "_
   &" LEFT JOIN kunder K on (kid = tknr) "_
   &" LEFT JOIN medarbejdere m ON (mid = tmnr) WHERE "& medarbSQLkri &" AND "_
   &" tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"' AND tfaktim = 5 ORDER BY tmnr, tdato DESC "
   
   'Response.Write strSQL
   'Response.flush
   
   at = 0
   timertot = 0
   lwedaynm = ""
   timerDag = 0
   ltmnr = 0
   lastMid = 0
   
   oRec.open strSQL, oConn, 3
   while not oRec.EOF 
    
    
    if lastMid <> oRec("mid") then
        call afsluger(oRec("mid"), sqlDatoStart, sqlDatoSlut)
    end if
      
    select case right(at, 1)
    case 0,2,4,6,8
    bgcol = "#ffffff"
    case else
    bgcol = "#EFF3ff"
    end select
    
    
    if ltmnr <> oRec("tmnr") then 
    if at <> 0 then%>
    <tr>
        <td align=right colspan=8><br />Total i periode: <b><%=formatnumber(kmTot, 2) %></b>
        <br />
        <%call ArTilDatoKm() %>
        
        </td>
   </tr>
   <%
   kmTot = 0
   end if %>
    <tr bgcolor="#cccccc">
        <td colspan=8 height=30><b><%=oRec("tmnavn") %></b> (<%=oRec("mnr") %>)
        <%if len(trim(oRec("init"))) <> 0 then
        %>
        - <%=oRec("init") %>
        <%
        end if %>
        </td>
    </tr>
    <%
    end if
    
    %>
   <tr bgcolor="<%=bgcol%>">
   <td valign=top style="width:100px; border-bottom:1px #cccccc solid; white-space:nowrap;"><b><%=left(weekdayname(weekday(oRec("Tdato"))), 3) %>. <%=day(oRec("Tdato")) &" "&left(monthname(month(oRec("Tdato"))), 3) &". "& right(year(oRec("Tdato")), 2)%></b></td>
   <td valign=top style="width:140px; border-bottom:1px #cccccc solid; white-space:nowrap;"><%=oRec("tmnavn") %> (<%=oRec("mnr") %>)
        <%if len(trim(oRec("init"))) <> 0 then
        %>
        - <%=oRec("init") %>
        <%
        end if 
        
        if len(trim(oRec("timerkom"))) <> 0 then
        timerKom = replace(oRec("timerkom"), "Til:", "<br><b>Til:</b> ")
        timerKom = replace(timerKom, "Fra:", "<b>Fra:</b> ")
        timerKom = replace(timerKom, "Via:", "<br><b>Via:</b> ")
        else
        timerKom = oRec("timerkom")
        end if
        
        %></td>
   <td valign=top style="border-bottom:1px #cccccc solid; width:200px"><b><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</b><br />
   <%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
   <td valign=top style="border-bottom:1px #cccccc solid; padding-right:20px; height:50px;"><%=oRec("taktivitetnavn") %></td>
    <td valign=top style="border-bottom:1px #cccccc solid; width:250px;"><%=timerKom %>&nbsp;</td>
    <td valign=top align=center style="border-bottom:1px #cccccc solid;"><%if oRec("bopal") = 1 then %>
    Ja
    <%else %>
    &nbsp;
    <%end if %></td>
    
    <%
    strEksportTxt = strEksportTxt & oRec("Tdato") & ";"& oRec("tmnavn") &" ("& oRec("mnr") &");" & oRec("tknavn") & " ("& oRec("kkundenr") &");"
    strEksportTxt = strEksportTxt & oRec("tjobnavn") &" ("& oRec("tjobnr")  &");" & oRec("taktivitetnavn") &";"& oRec("bopal") & ";" 
    %>
    
     <td valign=top  style="border-bottom:1px #cccccc solid;">
     
    
     
     <%
     
     lastDestNavn = ""
     
     call destAdr(oRec("destination")) %>
     
    <%=lastDestNavn%>&nbsp;</td>
   <td valign=top align=right style="border-bottom:1px #cccccc solid;">
          <%
        '** er periode godkendt ***'
		        tjkDag = oRec("Tdato")
		        erugeafsluttet = instr(afslUgerMedab(oRec("tmnr")), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                

                strMrd_sm = datepart("m", tjkDag, 2, 2)
                strAar_sm = datepart("yyyy", tjkDag, 2, 2)
                strWeek = datepart("ww", tjkDag, 2, 2)
                strAar = datepart("yyyy", tjkDag, 2, 2)

                if cint(SmiWeekOrMonth) = 0 then
                usePeriod = strWeek
                useYear = strAar
                else
                usePeriod = strMrd_sm
                useYear = strAar_sm
                end if

                
                call erugeAfslutte(useYear, usePeriod, oRec("tmnr"))
		        
		        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		        'Response.Write "tjkDag: "& tjkDag & "<br>"
		        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		         strSQL2 = "SELECT risiko FROM job WHERE jobnr = '"& oRec("Tjobnr") &"'"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						risiko = oRec2("risiko")
						end if
						oRec2.close

                 call lonKorsel_lukketPer(tjkDag, risiko)
              
		         
                'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 if ( ( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 ) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) = year(now) AND DatePart("m", tjkDag) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                
                ugeerAfsl_og_autogk_smil = 1
                else
                
                ugeerAfsl_og_autogk_smil = 0
                end if 
                
                if oRec("godkendtstatus") = 1 then
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = ugeerAfsl_og_autogk_smil
                end if
                
                
                if ugeerAfsl_og_autogk_smil = 0 OR (level = 10) then%>
                
                <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=650,height=500,resizable=yes,scrollbars=yes')" class=vmenu><%=formatnumber(oRec("timer"), 2) %></a>
	            </td>
                <%else %>
                <%=formatnumber(oRec("timer"), 2) %>
                <%end if %>
   </tr>
   
   <%
   call htmlparseCSV(timerKom)
   
   strEksportTxt = strEksportTxt & lastDestnavn & ";" & formatnumber(oRec("timer"), 2) & ";"" " & htmlparseCSVtxt & " "";xx99123sy#z"
   
   lwedaynm = weekdayname(weekday(oRec("Tdato")))
   ltmnr = oRec("tmnr")
   
   kmTot = kmTot + oRec("timer")
    
   lastMid = oRec("mid") 
    
   at = at + 1
   oRec.movenext
   wend
   oRec.close
   
   
   if at <> 0 then%>
   
   
   <tr>
        <td align=right colspan=8><br />Total i periode: <b><%=formatnumber(kmTot, 2) %></b>
         <br />
        <%call ArTilDatoKm() %></td>
   </tr>
   <tr>
  
   <%else %>
    <tr>
        <td colspan=8>Ingen registreringer i valgte uge.</td>
    </tr>
   <%end if %>
   </table>
   
   
   <%else 
   
   
    call akttyper2009(2)
    
    dim overligger, rest
    dim lastEndDato, dest_idnum
    dim dest_id, dest_days, days_all
    redim dest_id(400,1), dest_days(720,1), days_all(720,1)
    redim lastEndDato(400,1), dest_idnum(400,1)
    redim overligger(400,1), rest(400,m)
    
    sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
    datoSlut = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
    datoStart = dateadd("m", -24, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut)
    sqlDatoStart = year(datoStart) &"/"& month(datoStart) &"/"& day(datoStart)
    
    %>
    <br />
    <b>Periode: <%=formatdatetime(datoStart, 1)%> til <%=formatdatetime(datoSlut, 1)%>  </b>
   <br /><br />
   
   <table cellpadding=4 cellspacing=0 border=0 width=100%>
   
   <%
   t = 1
   if t = 1000 then
   %>
   <tr><td>
   
   Denne statistik er under udarbejdelse og forventes klar onsdag d. 1/7-2009
   
   <br /><br />
   Med venlig hilsen
   <br />
   OutZourCE Dev. Team.
   
   </td></tr>
   
   <%
   else
   
   %>
    
	   <tr bgcolor="#5582d2">
	   <td class=alt style="width:120px;"><b>Medarbejder</b></td>
	   <td class=alt style="width:180px;"><b>Kørsel til destination</b></td>
	   <td class=alt style="width:150px;"><b>Periode</b></td>
	    <td class=alt align=right style="padding-right:20px;"><b>Antal dage med<br /> kørsel i periode</b></td>
	   <td class=alt align=right style="padding-right:20px;"><b>Antal dage med <br />skattefri godtgørelse</b></td>
	   <td class=alt align=right style="padding-right:20px;"><b>Akkumuleret pr. destination</b><br />
	   (foregående 24 md. periode<br />nulstillet af mellemlig.<br /> periode på min. arb. 60 dage)
	   </td>
	   <td class=alt align=right style="padding-right:20px;"><b>Resterende dage</b></td>
    </tr>
    <%
    
    strEksportTxt = "Medarbejder;Mnr;Init;Destination;Antal arb. dage;Antal dage m. godtgørelse;Akkumuleret pr. dest.;Rest dage;xx99123sy#z"  
    
   
    strSQL = "SELECT tknr, timer, tdato, tmnavn, tmnr, m.init, m.mnr, destination, tfaktim, bopal FROM timer AS t "_
    &" LEFT JOIN medarbejdere AS m ON (m.mid = t.tmnr)"_
    &" WHERE "& medarbSQLkri &" AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"' AND (tfaktim = 5 AND bopal = 1) ORDER BY tmnr, tdato, destination"  '(("& aty_sql_realhours &") AND tfaktim <> 5) OR 
    
    'Response.Write strSQL
    'Response.Flush
    
    lastMed = 0
    'lastDato = ""
    lastDest = ""
    lastTfaktim = 0
    lastBopal = 0
    lastDato = "2002-1-1"
    'lastEndDatoPrev = lastDato
    
    x = 0
    y = 0
    z = 0
    m = 0
    destIds = 0
    
    isDestiWrt = ""
    antalDageMSkatteFri = 0
    
    antalDageDiff = 0
    
    oRec.open strSQL, oConn, 3
    while not oRec.EOF 
    
    
                '** Desiutnation angivet eller ej **'
                if len(trim(oRec("destination"))) <> 0 then
                thisDest = trim(oRec("destination"))
                else
                thisDest = "k_0"
                end if
    
    
    
     '** Navn overskrift ***'
    if z = 0 then
    %>
    <tr bgcolor="#8caae6">
        <td colspan=7 height=20>
        <b><%=oRec("tmnavn")%></b> (<%=oRec("mnr") %>)
        
        <%if len(trim(oRec("init"))) <> 0 then %>
        - <%=oRec("init") %>
        <%end if %>
        
        </td>
    </tr>
    <%
    
    '** Ny perstDato ***'
    perStDato = oRec("tdato")
    'lastEndDato(destIds,0) = perStDato
    end if    
    
    
                
    
    
    
    if z <> 0 then
            
            
            '** tjekker seneste indtasting, er det mere end 60 dage siden??
            sidsteIndtast = dateDiff("d",LastDato, oRec("tdato"),2,2)
            'Response.Write "sidsteIndtast" & sidsteIndtast
    
            if lastMed <> oRec("tmnr") OR _
            (lastDest <> thisDest AND ((cint(lastBopal) = 1 AND cint(lastTfaktim) = 5))) OR sidsteIndtast >= 60 then 
            
            
            call destAdr(lastDest)
                    
                   
            
            
                  
            
               
            select case right(y, 1)
            case 0,2,4,6,8
            bgcolor = "#eff3FF"
            case else
            bgcolor = "#ffffff"
            end select
            
             overliggerThis = 60
             antalDageDiff = 0
             
             tjkStDato = "2002-01-02"
             tjkSlDato = "2002-01-03"
             
             call beregnDage(m,x,lastDest,overliggerThis)
             call medarbDestLinie(m)
             
              '** Ny perstDato ***'
             perStDato = oRec("tdato")
             'korselsRecFundetpaDato = 1
            
             
             y = y + 1
             
             
            
            end if 'medarbid  / dest
            
            
              
    
    
    '** Navn overskrift ***'
    if lastMed <> oRec("tmnr") then
    
    call medarbDageiPer()
    
    %>
    
     
    
    <tr bgcolor="#8caae6">
        <td colspan=7 height=20>
        <b><%=oRec("tmnavn")%></b> (<%=oRec("mnr") %>)
        
        <%if len(trim(oRec("init"))) <> 0 then %>
        - <%=oRec("init") %>
        <%end if %>
        
        </td>
    </tr>
    <%
    
    antalArbIalt = 0
    end if   
    
    
     '** nulstiller ***'
    if lastMed <> oRec("tmnr") then
    m = m + 1
    redim overligger(400,m), dest_idnum(400,m), rest(400,m)
    redim lastEndDato(400,m), dest_id(400,m), dest_days(720,m), days_all(720,m)
    isDestiWrt = ""
    destIds = 0
    x = 0
    y = 0
    end if
    
    
    end if 'z <> 0
    
    
                
    
    
                
                '** Er destination brugt før ellerer den ny? ***'    
                if instr(isDestiWrt, ",#"& trim(thisDest) &"#") = 0 then
                
                
                dest_id(y,m) = thisDest
               
                
                'Response.Write "<br>kid: "& dest_id(y,m) &" y: "& y 
                isDestiWrt = isDestiWrt & ",#"&trim(thisDest)&"#" 
                dest_days(y,m) = 0
                days_all(y,m) = 0
                dest_idnum(y,m) = destIds
                overligger(y,m) = 60
                destIds = destIds + 1 
                overliggerThis = overligger(y,m)
                rest(y,m) = overligger(y,m)
                
                first = 1 
                else
                first = 0 
               
                end if
                
                
            '*** Tæller dage immellem ***'
            'for d = 0 to UBOUND(dest_id)
                
            '    if dest_id(d,m) <> thisDest OR z = 0 then
            '    dest_days(d,m) = dest_days(d,m) + 1
            ''    'Response.Write "<br>kid: "& dest_id(d,m) &" orecDEST:"& oRec("destination") &" d: "& d &" m: "& m &" : "& dest_days(d,m) &""
             '   end if
                
            'next
    
    
    
    lastDest = thisDest
    lastBopal = oRec("bopal")
    lastTfaktim = oRec("tfaktim")
    lastmedNavn = oRec("tmnavn")
    lastMed = oRec("tmnr")
    lastInit = oRec("init")
    lastDato = oRec("tdato")
    lastMnr = oRec("mnr")
    'lastKid = oRec("tknr")
    
    
    
    x = x + 1
    z = z + 1
    
    
    oRec.movenext
    wend
    oRec.close
    
    
    if z <> 0 then
    
    select case right(y, 1)
    case 0,2,4,6,8
    bgcolor = "#eff3FF"
    case else
    bgcolor = "#ffffff"
    end select
   
    
    call destAdr(lastDest)
    call beregnDage(m,x,lastDest,overliggerThis)
    call medarbDestLinie(m)
    call medarbDageiPer()
    
    end if
    
    
    end if 't = 1000 %>
    
   
    </table>
    
    <%end if %>
	
	</div>
	
	
	
	
	
	

	<br /><br /><br /><br /><br /><br /><br />&nbsp;
	</div>
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	
	
	<%
	
	public tjkStDato, tjkSlDato, antalDageMSkatteFri, lastRest, antalDageDiff 
	public days_All_thisDest
	function beregnDage(m,x,lastDest,overliggerThis)
	
	 for d = 0 to UBOUND(dest_id)
        
                    if dest_id(d,m) = lastDest then
                    
                    '** Beregner dage ***'
                     antalDageMSkatteFri = x 'dest_days(d,m)
                             
                                if dest_days(d,m) < overliggerThis then
                                    
                                    days_All_thisDest = days_all(d,m) + antalDageMSkatteFri
                                                        
                                else
                                    
                                    days_All_thisDest = antalDageMSkatteFri
                                    
                                end if    
                    
                    
                                    '** Er vi stadigvæk under overligger ***'
                                    if days_All_thisDest < overliggerThis then
                                    days_All_thisDest = days_All_thisDest
                                    antalDageMSkatteFri = antalDageMSkatteFri
                                    else
                                    'Response.Write "her" & overliggerThis & "<br>"
                                    days_All_thisDest_diff = (days_All_thisDest - overliggerThis)
                                    days_All_thisDest = days_All_thisDest - days_All_thisDest_diff
                                    antalDageMSkatteFri = x - days_All_thisDest_diff 
                                    end if
                    
                    
                    'dest_days(d,m) = 0
                    days_all(d,m) = days_all(d,m) + antalDageMSkatteFri 
                    
                    
                    rest(d,m) = rest(d,m) - antalDageMSkatteFri
                    lastRest = rest(d,m)
                    
                    antalDageDiff = dateDiff("d",lastEndDato(dest_idnum(d,m),m),perStDato,2,2)
                    lastEndDuse = lastEndDato(dest_idnum(d,m),m)
                    
                    tjkStDato = year(lastEndDuse) &"-"& month(lastEndDuse) &"-"& day(lastEndDuse)
                    tjkSlDato = year(perStDato) &"-"& month(perStDato) &"-"& day(perStDato)
                    
                    lastEndDato(dest_idnum(d,m),m) = lastDato
                    
                    end if
             next
	        
	            
	            
	             if antalDageDiff >= 60 then
                    
                    
                    
                    '** Finder antal dage med registreringer siden sidste (skal der nulstilles?)
                    strSQL2 = "SELECT tid AS antal FROM timer AS t "_
                    &" WHERE tmnr = "& lastMed &" AND tdato BETWEEN '"& tjkStDato &"' AND '"& tjkSlDato &"' AND "_
                    &" (("& aty_sql_realhours &") OR (tfaktim = 5 AND destination <> '"& lastDest &"'))" _
                    &" GROUP BY tdato ORDER BY tdato"  '(("& aty_sql_realhours &") AND tfaktim <> 5) OR 
                    
                    'Response.Write strSQL2 & "<br><br>"
                    'Response.Flush
                    antalArbDage = 0
                    oRec2.open strSQL2, oConn, 3
                    while not oRec2.EOF 
                    
                    antalArbDage = antalArbDage + 1 'oRec2("antal")
                    
                    oRec2.movenext
                    wend 
                    oRec2.close
                    
                    
                         'if antalArbDage > 2 then
                         if antalArbDage >= 60 AND lastEndDuse <> "" AND (cint(first) <> 1) then
                         %>
                         <tr bgcolor="#FFFFe1">
                            <td colspan=7><br />
                            Vedr.: <b><%=lastDestNavn %></b><br />
                            Nulstilling af 60 dages regel (tildeler yderligere 60 dage) på destination.<br /><br />
                           
                            Der har været mere end 60 arbejdsdage mellem seneste registering på denne destination og denne registering,<br />
                            derfor udviddes antal dage med skattefrigodtgørelse på denne destination med yderligere 60 dage.
                            <br />
                            Seneste kørsel til destination var: <b><%=lastEndDuse %> (Diff: <%=antalDageDiff %> / <%=antalArbDage %> arb. dage)</b>
                            </td>
                         </tr>
                         <%
                         
                            for d = 0 to UBOUND(dest_id)
        
                                if dest_id(d,m) = lastDest then
                                '** tilføjer nye 60 dage da der har været 60 arb.dage imellem.
                                overligger(d,m) = 60
                                rest(d,m) = overligger(d,m) - antalDageMSkatteFri 
                                overliggerThis = overligger(d,m) 
                                lastRest = rest(d,m)
                                'nulstilling = 1
                                end if
                            
                            next
                         
                         end if
             
             
             end if
	        
	            
	        
	end function
	
	
	public antalArbIalt
	function medarbDestLinie(m)
	
	 
	
	
	
     
    %>
    <tr bgcolor="<%=bgcolor %>">
    <td style="border-bottom:1px #cccccc solid;">
    
   
    
    <%=lastmedNavn &" ("& lastMnr &") "%>
        
        <%if len(trim(lastInit)) <> 0 then %>
        - <%=lastInit%>
        <%end if %>
        
       
    </td>
    
    <%          
    strEksportTxt = strEksportTxt & lastmedNavn &";"& lastMnr &";"& lastInit &";" 
    
   
               
    
                
               
               
    %>
    
    <td style="border-bottom:1px #cccccc solid; white-space:nowrap;"><%=lastDestNavnTxt%></td>
    <td style="border-bottom:1px #cccccc solid; white-space:nowrap;"><%=formatdatetime(perStDato, 1) %> - <%=formatdatetime(LastDato, 1) %></td>
    
    <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=formatnumber(x, 2) %></td>
    <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=formatnumber(antalDageMSkatteFri, 2) %></td>
    <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><b><%=formatnumber(days_All_thisDest, 2) %></b></td>
    <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=formatnumber(lastRest, 2) %></td>
    
    <%
    strEksportTxt = strEksportTxt & lastDestNavn &";"& formatnumber(x, 2) &";"& formatnumber(antalDageMSkatteFri, 2) &";"& formatnumber(days_All_thisDest, 2)  &";"& formatnumber(lastRest, 2) &";xx99123sy#z" 
     %>
    </tr>
    
    <%
    lastDestNavn = ""
    antalArbIalt = antalArbIalt + x 
    'antalDageMSkatteFri = 0
    isDestiWrt = isDestiWrt & ",#"&lastDest&"#" 
    'y = y + 1
    x = 0
    
    
    
    end function 
    
    
     'public lastDestNavnTxt
     'function thisDest(lastDest)
     '               call destAdr(lastDest)
     '               
     '               if len(trim(lastDestNavn)) <> 0 then
     '               lastDestNavnTxt = "<b>"&lastDestNavn&"</b>"
     '               lastDestNavn = lastDestNavn
     '               else
     '               lastDestNavnTxt = "<font class=""lillegray"">Destination ikke angivet</font>"
     '               lastDestNavn = "Destination ikke angivet"
     '               end if
     '               
    'end function
    
    public lastDestNavn, lastDestNavnTxt
    function destAdr(dest)
    if len(trim(dest)) <> 0 then
                        
                        
                        '** Er der kørt til en kontaktperson / filial eller til hovedkontakten.
                        '** KP = kontkapers
                        if left(dest,2) <> "kp" then
                            
                            if len(trim(dest)) > 2 then
                            lastDestId = 0
                            len_lastDest = len(dest)
                            lastDestId = right(dest, len_lastDest - 2)
                            else
                            lastDestId = 0
                            end if
                            
                            lastDestNavn = ""
                            strSQLk = "SELECT kkundenavn FROM kunder WHERE kid = "& lastDestId
                            
                            'Response.Write dest & " -sql: "& strSQLK & "<br>"
                            'Response.flush
                            
                            oRec2.open strSQlk, oConn, 3
                            
                            if not oRec2.EOF then
                            
                            lastDestNavn = oRec2("kkundenavn")
                            
                            end if
                            oRec2.close
                        
                        else
                        
                            lastDestId = 0
                            len_lastDest = len(dest)
                            lastDestId = right(dest, len_lastDest - 3)
                            
                            
                            lastDestNavn = ""
                            strSQLk = "SELECT navn FROM kontaktpers WHERE id = "& lastDestId
                            
                            'Response.write "kp tur: "& strSQLk & "<br>"
                            'Response.flush
                            oRec2.open strSQlk, oConn, 3
                            
                            if not oRec2.EOF then
                            
                            lastDestNavn = oRec2("navn")
                            
                            end if
                            oRec2.close
                            
                            end if 
                        
                        end if
    
    
    
                    if len(trim(lastDestNavn)) <> 0 then
                    lastDestNavnTxt = "<b>"&lastDestNavn&"</b>"
                    lastDestNavn = lastDestNavn
                    else
                    lastDestNavnTxt = "<font class=""lillegray"">Destination ikke angivet</font>"
                    lastDestNavn = "Destination ikke angivet"
                    end if
    
    end function
    
    
    function medarbDageiPer()
    %>
    <tr>
        <td></td>
        <td></td>
        <td></td>
        <td align=right style="padding-right:20px;"><b><%=formatnumber(antalArbIalt) %></b></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    <%
    
    end function
    %>
    
    
    
    <%if media = "export" then 
    

    call TimeOutVersion()
   
	ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stat_korsel.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
			    'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				'objF.WriteLine(strOskrifter & chr(013))
				objF.WriteLine(ekspTxt)
				
				
				
				
	
	Response.redirect "../inc/log/data/"& file &""	

end if %>
    
    
    
    <%if media <> "print" then

	
	            ptop = 102 
                pleft = 730
                pwdt = 130

                call eksportogprint(ptop,pleft, pwdt)
                %>
                
                
                    
                     <tr>
                    <td align=center>
                    <a href="#" onclick="Javascript:window.open('stat_korsel.asp?media=export&FM_sumprmed=<%=showtot %>&medarbSel=<%=medarbSel %>', '', 'width=100,height=100,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
                    </td><td><a href="#" onclick="Javascript:window.open('stat_korsel.asp?media=export&FM_sumprmed=<%=showtot %>&medarbSel=<%=medarbSel %>', '', 'width=100,height=100,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport</a>
                    </td>
                   </tr>
                   
                  
                    
                    <tr>
                    <td align=center>
                   <a href="stat_korsel.asp?media=print&FM_sumprmed=<%=showtot %>&medarbSel=<%=medarbSel %>" target="_blank">
                   &nbsp;<img src="../ill/printer3.png" border=0 alt="Print version" /></a>
                </td><td><a href="stat_korsel.asp?media=print&FM_sumprmed=<%=showtot %>&medarbSel=<%=medarbSel %>" target="_blank" class=rmenu>Print version</a></td>
                   </tr>
                  
                    </table>
                      </div>
                  
                  
                  
               
                   
           
                <%else%>

                <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %>
                <%end if%>
	
	
	
	
	
	
	
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
