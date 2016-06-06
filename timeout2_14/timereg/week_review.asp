<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "week_review.asp"
	
	select case func 
	case "-"
	
	case else
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	leftPos = 20
	topPos = 132
	

	if len(request("FM_md")) <> 0 then
	strMd = request("FM_md")
	strAar = request("FM_aar")
	intMid = request("FM_medarb")
	else
	intMid = session("mid")
	strMd = month(date)
	strAar = year(date)
	end if
	
	
	
	if len(request("FM_md_slut")) <> 0 then
	strMd_slut = request("FM_md_slut")
	strAar_slut = request("FM_aar_slut")
    else
	strMd_slut = datepart("m",(dateadd("m", 1, date)))
	    
	    if strMd = 12 then
	    strMd_slut = 1
	    strAar_slut = datepart("yyyy",(dateadd("yyyy", 1, date)))
	    else 
	    strAar_slut = datepart("yyyy",(date))
	    end if
	    
	end if
	
	
	select case strMd_slut
	case 1,3,5,7,8,10,12
	usedag = 31
	case 2
	    select case useMD_slut
	    case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040, 2044
	    usedag = 29
	    case else 
	    usedag = 28
	    end select
	case else
	usedag = 30
	end select
	
	
	startdato = strAar&"/"&strMd&"/1"
	slutdato = strAar_slut&"/"&strMd_slut&"/"&usedag
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<h3>Ugesedler</h3>
	
	<%call filterheader(0,0,600,pTxt)%>
	<table border=0 cellspacing=2 cellpadding=2>
	<form action="week_review.asp" method="post" name="periode" id="periode">
	<tr>
	<td valign=top width=200><b>Medarbejder:</b><br /> 
	<select name="FM_medarb" id="FM_medarb">
	<option value="0">Alle</option>
	<%
	strSQL = "SELECT mnavn, mid, init, mnr FROM medarbejdere WHERE mansat <> '2' AND mansat <> '3' ORDER BY mnavn"
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	if len(request("FM_medarb")) = 0 AND x = 0 then
		selthis = "SELECTED"
		intMid = session("mid") 'oRec("mid")
	else	
		if cint(intMid) = oRec("mid") then
		selthis = "SELECTED"
		else
		selthis = ""
		end if
	end if
	%>
	<option value="<%=oRec("mid")%>" <%=selthis%>><%=oRec("mnavn")%> (<%=oRec("mnr") %>) <%=oRec("init") %></option>	
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	
	%></select>
	</td>
	
	<td valign=top><b>Fra:</b><br />
	<select name="FM_md" id="FM_md">
	<%
	x = 1
	for x = 1 to 12
	select case x
	case 1
	strMDname = "1. Januar"
	case 2
	strMDname = "1. Februar"
	case 3
	strMDname = "1. Marts"
	case 4
	strMDname = "1. April"
	case 5
	strMDname = "1. Maj"
	case 6
	strMDname = "1. Juni"
	case 7
	strMDname = "1. Juli"
	case 8
	strMDname = "1. August"
	case 9
	strMDname = "1. September"
	case 10
	strMDname = "1. Oktober"
	case 11
	strMDname = "1. November"
	case 12
	strMDname = "1. December"
	end select
	
	if cint(strMd) = x then
	selthis = "SELECTED"
	showstrMDname = strMDname
	else
	selthis = ""
	end if
	%>
	<option value="<%=x%>" <%=selthis%>><%=strMDname%></option>
	<%next%>
	</select></td><td valign=top>
	<br />
	<select name="FM_aar" id="FM_aar">
	<%
	for x = 0 to 20
	
	strAarnameid = (2001 + x)
	
	if cint(strAar) = strAarnameid then
	selthis = "SELECTED"
	else
	selthis = ""
	end if
	%>
	<option value="<%=strAarnameid%>" <%=selthis%>><%=strAarnameid%></option>
	<%next%>
	</select></td>
	<td valign=top><b>Til: </b><br />
	<select name="FM_md_slut" id="FM_md_slut">
	<%
	x = 1
	for x = 1 to 12
	select case x
	case 1
	strMDname = "31. Januar"
	case 2
	strMDname = "(28). Februar"
	case 3
	strMDname = "31. Marts"
	case 4
	strMDname = "30. April"
	case 5
	strMDname = "31. Maj"
	case 6
	strMDname = "30. Juni"
	case 7
	strMDname = "31. Juli"
	case 8
	strMDname = "31. August"
	case 9
	strMDname = "30. September"
	case 10
	strMDname = "31. Oktober"
	case 11
	strMDname = "30. November"
	case 12
	strMDname = "31. December"
	end select
	
	if cint(strMd_slut) = x then
	selthis = "SELECTED"
	showstrMDname = strMDname
	else
	selthis = ""
	end if
	%>
	<option value="<%=x%>" <%=selthis%>><%=strMDname%></option>
	<%next%>
	</select></td><td valign=top>
	<br />
	<select name="FM_aar_slut" id="FM_aar_slut">
	<%
	for x = 0 to 20
	
	strAarnameid = (2001 + x)
	
    if cint(strAar_slut) = strAarnameid then
	selthis = "SELECTED"
	else
	selthis = ""
	end if
	%>
	<option value="<%=strAarnameid%>" <%=selthis%>><%=strAarnameid%></option>
	<%next%>
	</select></td>
	    
	    
	    <td valign=top style="padding-top:16px;">
	<input type="image" src="../ill/pilstorxp.gif">
	</td></tr>
	<tr>
	    <td>
            <input id="Checkbox1" type="checkbox" /> Vis enheder<br />
            <input id="Radio1" type="radio" /> Vis kun ..</td>
	</tr>
	</form></table>
    
    <!-- filter header sLut -->
	</td></tr></table>
	</div>
    
	
	
	
	
	<%
	
	
	
	
	'** Ugesedler Review / godkend ***'
	'*** Viser timeregistreringer for valgte uge ****'
       

    oimg = "ikon_timereg_48.png"
	oleft = 0
	otop =  30
	owdt = 500
	'oskrift = tsa_txt_064 &" "& datepart("ww", tjekdag(1), 2 ,2) & " "& datepart("yyyy", tjekdag(1), 2 ,3) 
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	
            
            if print <> "j" then
            tTop = 50
	        tLeft = 0
	        tWdth = 700
	        else
	        tTop = 100
	        tLeft = 40
	        tWdth = 700
	        end if
        	
        	
	        call tableDiv(tTop,tLeft,tWdth)
            %>
            <table cellpadding=2 cellspacing=1 border=0 id="inputTable" background-color="#d6dff5" width=100%>
	        <%if print <> "j" then %>
	        <tr>
	            <td colspan=4><a href="timereg_akt_2006.asp?showakt=1&fromsdsk=<%=fromsdsk%>" class=vmenu><%=tsa_txt_258%>..</td>
	        </tr>
	        <%end if%>
	        
	        <tr>
		        <td colspan=4 valign=top style="padding-top:5px;">
        		
        		
        
        		
        	
	        <%
	        
	        call ugesedler(startdato, slutdato, intMid)
	        
	        
	       
	       
	       
	        function ugesedler(startdato, slutdato, mid)
	        
	        
	        %>
	        <table cellpadding=4 cellspacing=0 border=0 width=100%>
	        <%
        	
	        '*** Aktiviteter og Timer SQL MAIN **** 
	        strSQL = "SELECT t.tid, t.taktivitetnavn, t.timer, t.tdato, t.timerkom, t.offentlig, tjobnavn, tjobnr, "_
	        &" k.kkundenavn, k.kkundenr, t.godkendtstatus, "_
	        &" t.tmnavn, j.id AS jid, t.tfaktim, t.timepris, t.valuta, v.valutakode "_
	        &" FROM timer t "_
	        &" LEFT JOIN job j ON j.jobnr = t.tjobnr "_
	        &" LEFT JOIN kunder k ON k.kid = j.jobknr "_
	        &" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
	        &" WHERE "_
	        &" t.tmnr = "& mid &" AND Tdato BETWEEN  '"& startdato &"' "_
	        &" AND '"& slutdato &"'"_
	        &" ORDER BY Tdato, t.tjobnavn, t.taktivitetnavn"
        	
	       

	        'Response.write strSQL
	        'Response.flush
	        
	        
	        lastdivval = 0
	        timerLastDato = 0
	        lastDato = ""
	        x = 0
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF
        		
		            if lastDato <> oRec("tdato") then%>
			        <%if x > 0 then%>
			        <tr>
				        <td colspan=5 align=right><%=tsa_txt_065 %>: <b><%=formatnumber(timerLastDato, 2)%></b></td>
			        </tr>
			        
			        
			        <%
			        timerugeTot = timerugeTot + timerLastDato
			        timerLastDato = 0
			        end if
			        
			        if lastweek <> datepart("ww", oRec("tdato"), 2 ,2) &" "& datepart("yyyy", oRec("tdato"),2,2) then 
        		    %>
        		    <tr>
        		    <td colspan=5 bgcolor="#FFC0CB" style="border:1px lightpink solid;">
        		    <h4 class="h4nobreak"><%=tsa_txt_005 %>: <%=datepart("ww", oRec("tdato"),2,2) &" "& datepart("yyyy", oRec("tdato"),2,2) %></h4>
        		    
        		    </td>
        		    </tr>
        		    
        		    <%
        		    end if
		       
		        %>
        		
		        
        		
        	
        		
        		
		        
        		
		        <%
		        diviswrt = diviswrt & ",#"& weekday(oRec("tdato")) &"#"
		        end if
        		
		        if lastDato <> oRec("tdato") then%>
		        <tr>
				        <td style="padding-top:20px;" colspan=5><b><%=weekdayname(weekday(oRec("tdato"))) &"&nbsp;"& formatdatetime(oRec("tdato"), 1)%></b></td>
		        </tr>
		        <tr bgcolor="#D6Dff5">
			        <td style="border-top:1px #8caae6 solid; border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_066 %></b></td>
			        <td style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_067 %></b></td>
			        <td style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_068 %></b> (<%=tsa_txt_069 %>)</td>
			        <td align=right style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_070 %></b></td>
			        <td align=right style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><b><%=tsa_txt_186 %></b></td>
			    </tr>
		        <%
        		
	            timerugeTot = timerugeTot + timerLastDato
	            'Response.Write timerugeTot & "<br>"
        		
		        end if%>
        		
		        <tr bgcolor="#ffffff">
		        <td valign=top style="border-bottom:1px #c4c4c4 dashed;">
		        <%if len(oRec("kkundenavn")) > 25 then%>
		        <%=left(oRec("kkundenavn"), 25)%>..
		        <%else%>
		        <%=oRec("kkundenavn")%>
		        <%end if%> 
		        (<%=oRec("kkundenr")%>)</td>
		        <td valign=top style="border-bottom:1px #c4c4c4 dashed;">
		        
		        <%if print <> "j" then%>
		        <a href="jobs.asp?menu=job&func=red&id=<%=oRec("jid")%>&int=1&rdir=treg" class=vmenu target="_top">
		        <%end if %>
		        
		        <%if len(oRec("tjobnavn")) > 25 then%>
		        <%=left(oRec("tjobnavn"), 25)%>..
		        <%else%>
		        <%=oRec("tjobnavn")%>
		        <%end if%> 
	 	        (<%=oRec("tjobnr")%>)
	 	        
	 	        <%if oRec("godkendtstatus") = 0 AND maxl <> 0 AND print <> "j" then%>
	 	        </a>
	 	        <%end if%>
	 	        
	 	        </td>
		        <td valign=top style="border-bottom:1px #c4c4c4 dashed;"><b><%=oRec("taktivitetnavn")%></b>
        		
		        <%
		        call akttyper(oRec("tfaktim"), 4) 
		        Response.Write "&nbsp;<font class=megetlillesort>("& akttypenavn &")"%></td>
		        <td align=right valign=top style="border-bottom:1px #c4c4c4 dashed;">
        		
		        <%
		        '******* Skal det være muligt at redigere indtastning ****
		        '*** Lastfakdato
		        lastfakdato = "1/1/2001"
        		
		        strSQLFAK = "SELECT f.fakdato FROM fakturaer f WHERE f.jobid = "& oRec("jid") &" AND f.fakdato >= '"& varTjDatoUS_man &"' AND faktype = 0 ORDER BY f.fakdato DESC"
		        oRec2.open strSQLFAK, oConn, 3
		        if not oRec2.EOF then
			        if len(trim(oRec2("fakdato"))) <> 0 then
			        lastfakdato = oRec2("fakdato")
			        end if
		        end if
		        oRec2.close
        		
        		'*** Er uge godkendt ***'
		        'call fakfarver(lastfakdato, tjekdag(1), oRec("tdato"))
        		
		        'if maxl <> 0 then
		        if oRec("godkendtstatus") = 0 AND maxl <> 0 AND print <> "j" then%>
		        <!--a href="rediger_tastede_dage.asp?id=<=oRec("Tid")%>&medarb=<=oRec("Tmnavn")%>&jobnr=<=intJobnr%>&eks=<=request("eks")%>&lastFakdag=<=lastFakdag%>&selmedarb=<=selmedarb%>&selaktid=<=selaktid%>&FM_job=<=request("FM_job")%>&FM_medarb=<=request("FM_medarb")%>&FM_start_dag=<=strDag%>&FM_start_mrd=%=strMrd%>&FM_start_aar=<=strAar%>&FM_slut_dag=<=strDag_slut%>&FM_slut_mrd=<=strMrd_slut%>&FM_slut_aar=<=strAar_slut%>" class="vmenu">-->
		        <a href="javascript:popUp('rediger_tastede_dage_2006.asp?id=<%=oRec("Tid")%>','600','500','250','120');" target="_self" class=vmenu>
		        <%end if%>
        		
		        <%=formatnumber(oRec("timer"), 2)%>
        		
		        <%if oRec("godkendtstatus") = 0 AND maxl <> 0 AND print <> "j" then%>
		        </a>
		        <%end if%></td>
		        <td style="border-bottom:1px #c4c4c4 dashed;" align=right><%=formatnumber(oRec("timepris"), 2) &" "& oRec("valutakode") %></td>
        		
	        </tr>
	        <%
	        x = x + 1
	        select case oRec("tfaktim")
	        case 1,2,6,13,14,20,21
	        timerLastDato = timerLastDato + oRec("timer")
            case else
	        timerLastDato = timerLastDato
	        end select
        	
	        lastDato = oRec("tdato") 
	        lastWeek = datepart("ww", oRec("tdato"),2,2) &" "& datepart("yyyy", oRec("tdato"),2,2)
        	
	        oRec.movenext
	        wend
        	
	        oRec.close

	        timerugeTot = timerugeTot + timerLastDato
        	
	        %>
        	
	        <%if x > 0 then%>
	        <tr>
		        <td colspan=5 align=right><%=tsa_txt_065 %>: <b><%=formatnumber(timerLastDato, 2)%></b></td>
	        </tr>
	        <tr bgcolor="#FFDFDF">
		        <td colspan=5 align=right><%=tsa_txt_065 %>&nbsp;<%=tsa_txt_071 %>&nbsp;<%= datepart("ww", lastDato, 2,2)%>, <%= datepart("yyyy", lastDato, 2 ,2)%>: <b><%=formatnumber(timerugeTot, 2)%></b></td>
	        </tr>
	        
	        
	        
	        
	        
	        
	                <%if print <> "j" then %>
	                <br />
	                <a href="timereg_akt_2006.asp?print=j" target="_blank" class=vmenu><%=tsa_txt_072 %>..</a> 
	                <%end if%>
                	
            <!--</div>-->
            <%else %>
            <tr>
                <td colspan=5 style="padding-top:20px;"><b><%=tsa_txt_124 %></b></td>
            </tr>
            
            <%end if%>
            
            </table>
        	
        	
	       <%end function %>
        	
        	
	       
	        <br /><br /><br />
                &nbsp;
	
	
    
    </td>
    </tr>
            </table>
    
    
    <!-- table div --> 
	</div>
	
	
	
	
	
	
	<br /><br />
	  <% 
       itop = 40
       ileft = 0 
       iwdt = 400
       call sideinfo(itop,ileft,iwdt)%>		
       
   
		
	    <br />Insfdfsdf
	    
                &nbsp;
	   
	    <!-- side info slut -->
        </td></tr></table>
			</div>
	    
	
	<br /><br /><br /><br /><br /><br /><br /><br />
        &nbsp;
			
			</div>
	
	<% 
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
