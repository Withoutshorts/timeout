<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->





<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
    menu = request("menu")
    thisfile = "timerimperr"

    media = request("media")
    if media = "print" OR media = "export" then
        print = "j"
    end if

	
      if len(trim(request("FM_sog"))) <> 0 then
   sogTxt = request("FM_sog")
   else
   sogTxt = ""
   end if


   orgSEL0 = ""
   orgSEL1 = ""

   if len(trim(request("FM_origin"))) <> 0 then
   origin = request("FM_origin")
        if origin = 1 then 
        orgSEL1 = "SELECTED"
        tfaktimUse = 1'** Skal rettes **'
        originTXT = "Vietnam"
        else
   
        origin = 3
       orgSEL0 = "SELECTED"
       tfaktimUse = 91
       originTXT = "CATI"
        end if
   else
   origin = 3
   orgSEL0 = "SELECTED"
   tfaktimUse = 91
   originTXT = "CATI"
   end if
    
    
    func = request("func")
	
	

    select case func
    case "flyt"
        
        if len(trim(request("FM_oldjobnr"))) <> 0 then
        oldJobnr = trim(request("FM_oldjobnr"))
        else
        oldJobnr = "0"
        end if


        if len(trim(request("FM_jobnr"))) <> 0 then
        newJobnr = trim(request("FM_jobnr"))
        else
        newJobnr = "0"
        end if

    

        if oldJobnr <> "0" AND newJobnr <> "0" then
        strSQLjob = "SELECT j.id, j.jobnavn, j.serviceaft, j.jobknr, j.jobnr, k.kkundenavn, k.kkundenr, k.kid, a.navn AS aktnavn, "_
        &" a.id AS aid FROM job j "_
        &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
        &" LEFT JOIN aktiviteter a ON (a.navn = 'Epinion Asia (Vietnam) time' AND a.job = j.id) WHERE jobnr = '"& newJobnr &"'"
        'Response.Write strSQLjob
        'Response.flush

        oRec.open strSQLjob, oConn, 3
        if not oRec.EOF then

            '** Opdater **
            if isNull(oRec("aktnavn")) = false then

            strSQLupd = "UPDATE timer SET tjobnavn = '"& oRec("jobnavn") &"', tjobnr = "& oRec("jobnr") & ", tknavn = '"& oRec("kkundenavn") &"', "_
            &" tknr = "& oRec("kid") &", seraft = "& oRec("serviceaft") &", taktivitetid = "& oRec("aid") &" WHERE timerkom = '"& oldJobnr &"' AND tjobnr = 100"


            'Response.Write strSQLupd
            'Response.flush
            oConn.execute(strSQLupd)

            end if

            

        end if
        oRec.close
        
        end if
    
    'Response.end
    Response.redirect "timerimperr.asp?FM_sog="&sogTxt&"&FM_origin="&origin

    end select
	
	
	
	
	
	
	
	'************ slut faste filter var **************		
	

	if media = "print" OR media = "export" then
	'leftPos = 20
	'topPos = 62
    %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	else
	'leftPos = 90
	'topPos = 102
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    

	
	<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%
        'call tsamainmenu(7)
    %>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
        'call stattopmenu()
    %>
	</div>
	<%
	end if
	%>-->


	<%

    call menu_2014()

	pleft = 90
	ptop = 82
	'ptopgrafik = 348

    tkMinusDage = 40 '31 '65

 	%>	
	<div id="Div1" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible;">
	
	
	<%
	
	

   


	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	

  
    tTop = 20
	tLeft = 0
	tWdth = 620
	
	
	call tableDiv(tTop,tLeft,tWdth)



        'oimg = "icon_band_aid.png"
	oleft = 0
	otop = 0
	owdt = 400
	oskrift = "Timer Import Errors (sen. "& tkMinusDage &" dage +)"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)

    %>

    <br />
    <b>Fejlforklaring:</b><br /><br />
    1: MedarbejderId mangler i import data <br />
    11: MedarbejderId ikke fundet i TimeOut <br />
    2: Jobnr mangler i import data<br />
    21: Job ikke fundet i TimeOut <br />
    3: Aktivitet ikke fundet på job <br />
    4: Valuta ikke fundet <br />
    5: Kunde ikke fundet på job<br />
    51: Kundenavn ikke fundet på valgte kunde<br />
    6: Externt sysid allerede indlæst<br />
  
    7: Timer er i et ukendt format / mangler i import data<br />
    8: Dato er i et ukendt format / mangler i import data<br />
    9: Uge er lukket<br />
    13: Bonus Akt. ikke fundet på job<br /><br />
    
    <%
    ddminusSeven = dateadd("d", -tkMinusDage, now) 
    ddminusSevenTXT = ddminusSeven
    ddminusSeven = year(ddminusSeven) & "-" & month(ddminusSeven) & "-"& day(ddminusSeven)
    dd = year(now) &"-"& month(now) &"-"& day(now) 
    ddTxt = day(now) &"-"& month(now) &"-"& year(now) 
     %>


    Periode: <h3><%=formatdatetime(ddminusSevenTxt,1)%> - <%=formatdatetime(ddTxt, 1)%>   </h3>     
   

  

   <form method="post" action="timerimperr.asp">
   Origin: <select id="Select1" name="FM_origin" onChange="submit();">
            <option value="3" <%=orgSEL0%>>CATI</option>
            <option value="1" <%=orgSEL1%>>Vietnam</option>
        </select><br /><br />
   Søg på Id, Eksternt id, Job nr. eller Intw nr.: <input id="FM_sog" name="FM_sog" value="<%=sogTxt%>" type="text" />&nbsp;<input
       id="Submit1" type="submit" value=" Søg >> " /><br />
   % = Wildcard, ved søgning vises d.d. - <%=tkMinusDage%> dage.
   
   </form>
        

   <table cellspacing=1 cellpadding=2 border=0 width=100%>
   <%
   
   totalAllSeven = 0

    if sogTxt <> "" then
    ub = 0 
    ddminusSevenUseTxt = dateAdd("d", -tkMinusDage, ddTxt)
     ddminusSevenUse = year(ddminusSevenUseTxt) &"/"& month(ddminusSevenUseTxt) &"/"& day(ddminusSevenUseTxt)
    else
    ub = tkMinusDage
    end if

   for days = 0 to ub
   
   if sogTxt <> "" then
    ddminusSevenUseTxt = dateAdd("d", -tkMinusDage, ddTxt)
     ddminusSevenUse = year(ddminusSevenUseTxt) &"/"& month(ddminusSevenUseTxt) &"/"& day(ddminusSevenUseTxt)
    else
   ddminusSevenUseTxt = dateAdd("d", -days, ddTxt)
   ddminusSevenUse = year(ddminusSevenUseTxt) &"/"& month(ddminusSevenUseTxt) &"/"& day(ddminusSevenUseTxt)
   end if
   %>
   
   <tr><td colspan=7><br /><br />
   <%if sogTxt <> "" then %>
   Timeregdato >=
   <%else %>
   Timeregdato:
   <%end if %>

    <b><%=weekdayname(weekday(ddminusSevenUseTxt)) &" "& formatdatetime(ddminusSevenUseTxt, 1)%></b>
    
    <!-- tjekker dagen er indlæst i timereg -->
    <%
    strSQLtjkindl = "SELECT COUNT(tid) AS antalRec FROM timer WHERE tfaktim = "& tfaktimUse &" AND tdato = '"& ddminusSevenUse &"' AND extsysid <> 0 AND origin = "& origin &" GROUP BY tdato"

    'Response.Write strSQLtjkindl
    'Response.flush

    oRec.open strSQLtjkindl, oConn, 3
    antalRecIdag = 0
    If not oRec.EOF then
    antalRecIdag = oRec("antalRec")
    end if
    oRec.close
    %>
    
    <br />Antal <%=originTxt %> registreringer indlæst i TimeOut: <b><%=antalRecIdag%></b>
    </td></tr>
    <tr bgcolor="#5582d2">
        <td class=alt>Id</td>
        <td class=alt>Kørselsdato</td>
        <td class=alt>Externt ID</td>
        <td class=alt>Error ID</td>
        <td class=alt>Job Nr.</td>
        <td class=alt>Job ID.</td>
        <td class=alt>Intw nr.</td>
        <td class=alt>Timereg. Dato</td>
        <td class=alt>Timer (1,00 = 1 time)</td>
    </tr>
    
    
    <%
    'dd' = ypar()
    
    if sogTxt <> "" then
    sogfSQL = " AND (id LIKE '"& sogTxt &"' OR extsysid LIKE '"& sogTxt &"' OR med_init LIKE '"& sogTxt &"' OR jobnr LIKE '"& sogTxt &"') AND timeregdato >= '"& ddminusSevenUse &"'"
    else
    sogfSQL = " AND timeregdato = '"& ddminusSevenUse &"'"
    end if
   

    strSQLtimerimperr = "SELECT id, dato, extsysid, errid, jobid, jobnr, med_init, timeregdato, timer FROM timer_imp_err WHERE errid  <> 6 "& sogfSQL &" AND extsysid <> 0 AND origin = "& origin &" GROUP BY extsysid ORDER BY timeregdato, dato, errid"
    
    'Response.Write strSQLtimerimperr & "<br><br>"
    'Response.flush

    oRec.open strSQLtimerimperr, oConn, 3
    x = 0
    
    while not oRec.EOF 

    select case right(x, 1)
    case 0,2,4,6,8
    bgthis = "#8caae6"
    case else
    bgthis = "#FFFFFF"
    end select
        %>
    <tr bgcolor="<%=bgthis %>">
        
        <td><%=oRec("id") %></td>
        <td><%=oRec("dato") %></td>
        <td><%=oRec("extsysid") %></td>
        <td><%=oRec("errid") %></td>
        <td><%=oRec("jobnr") %></td>
        <td><%=oRec("jobid") %></td>
        <td><%=oRec("med_init") %></td>
        <td><%=oRec("timeregdato") %></td>
        <td><%=oRec("timer") %></td>
    </tr>


    <%
    x = x + 1
    oRec.movenext
    wend
    oRec.close

    %>
    <tr>
    <td colspan=8>Antal ialt: <%=x %></td>
    
    </tr>

    <%
    totalAllSeven = totalAllSeven + x
    next %>
	

    <tr>
    <td colspan=8>Antal ialt 7 dage: <%=totalAllSeven %></td>
    
    </tr>
    </table>



	
     </div><!-- table div -->

     </div><!-- side div -->


     <%if origin = "1" then %>
     <div id="Div2" style="position:absolute; left:700px; top:200px; visibility:visible;">
	
	<%
    
     tTop = 20
	tLeft = 20
	tWdth = 620
	
	
	call tableDiv(tTop,tLeft,tWdth)
     %>
     <table cellpadding=0 cellspacing=0 border=0 width=100%>
     <tr><td><br />
     <h3>Timer på Epinion Asia</h3>
     <b>Diverse Vietnam projekter (lokale/herreløse) (100)</b>


     <table cellpadding=2 cellspacing=0 border=0 width=100%>
     
     
     <%
     j = 0
     strSQL = "SELECT sum(timer) AS timer, tjobnavn, tjobnr, timerkom, tdato FROM timer WHERE tdato >= '2010-01-01' AND tjobnr = 100 AND timerkom <> '' GROUP BY timerkom ORDER BY tdato DESC"
     oRec.open strSQL, oConn, 3
     While not oRec.EOF
        
        
        select case right(j, 1)
        case 0,2,4,6,8
        bgthis = "#eff3ff"
        case else
        bgthis = "#ffffff"
        end select 
    
     %>
     <form id="<%=j%>" action="timerimperr.asp?func=flyt&FM_origin=<%=origin %>&FM_sog=<%=sogTxt%>" method="post">
     <tr bgcolor="<%=bgthis%>"><td style="width:300px;"><%="Projekt: "& oRec("timerkom") &"- dato: "& oRec("tdato") &" - timer: " & oRec("timer") %></td>
     <td>flyt til jobnr: </td>
     <td><input id="Text2" type="hidden" name="FM_oldjobnr" value="<%=trim(oRec("timerkom") ) %>" /><input id="Text1" name="FM_jobnr" type="text" style="font-size:9px; width:80px;" /></td>
     <td><input id="Submit2" type="submit" value="flyt >> " style="font-size:9px;" /></td></tr>
     </form>    
     
     <%

     j = j + 1
     Response.flush
     oRec.movenext
     wend
     oRec.close
       %>

       </table>
     </td></tr>
     
     </table>

    </div><!-- table div -->

	<br><br>
	<br>
	
 

	<br>
	<br>
	
	&nbsp;
	
	
	</div>
    <%end if %>

    <br>
	<br>
	
	&nbsp;
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
