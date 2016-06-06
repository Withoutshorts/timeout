<%response.buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
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
	
    '********************************************************************************************************
    '**** EARLY BINDING VARIABLES ***
    '********************************************************************************************************


    level = session("rettigheder")

    media = request("media")
	func = request("func")
	if len(trim(request("yuse"))) <> 0 then
	ysel = request("yuse")
	else
	ysel = year(now)
	end if
	
	if len(trim(request("FM_start_mrd"))) <> 0 then
	strMrd = request("FM_start_mrd")
	else
        select case lto
        case "adra"
        strMrd = month(now)
        case else 
	    strMrd = 1
        end select
	end if
	
    select case request("per_interval") 
      
      case 1
      per_i_enCHK = "SELECTED"
      per_interval = 1
      daysInInterval = 31
      faktor = 4
      weekspan = 7

      case 3 
      per_i_treCHK = "SELECTED"
      per_interval = 3
      daysInInterval = 170
      faktor = 4
      weekspan = 7
      case else

        select case lto
        case "adra"

          per_i_treCHK = "SELECTED"
          per_interval = 3
          daysInInterval = 170
          faktor = 4
          weekspan = 7

        case else

          per_i_tolvCHK = "SELECTED" 
          per_interval = 12 
          daysInInterval = 353
          faktor = 1
          weekspan = 1
        
        end select

      end select
	
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
	thisfile = "feriekalender"

    if len(trim(request("vis"))) <> 0 then
    vis = request("vis")
    else
    vis = 0
    end if
	
                select case vis
                case 1
				strOskrifter = "Medarbejder;Nr.;Init;Ferie optjent (incl. overført og opt. u. løn);Måned;Ferie afholdt;Ferie udbetalt;Ferie afholdt u. løn;Ferie saldo;;Feriefridage Optjent;Feriefridage afholdt; Feriefridage udbetalt;Feriefridage saldo;"
                case 2
				strOskrifter = "Medarbejder;Nr.;Init;Ferie optjent (incl. overført og opt. u. løn);Ferie afholdt + udb.;Ferie afholdt u. løn;Ferie saldo;Ferie + Feriefridage planlagt;Sygdom; Barn syg; Afspadsering;Feriefridage optjent;Feriefridage afholdt + Udb.;Feriefridage saldo;"
                case else
                strOskrifter = "Medarbejder;Nr.;Init;Dato;Type;Timer;Dage;"
                end select

	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	media = request("media")

	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then
	        if media <> "print" then
	        thisMiduse = request("FM_medarb_hidden")
    	    else
    	    thisMiduse = request("FM_medarb")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
	
	
     '**** END EARLY BINDING VARIABLES ***
    
	
	



    '********************************************************************************************************
	'**** PRINT VERSION HEADER OR SCREEN HEADER ***
    '********************************************************************************************************
	
    if media = "print" then
            %>
            <html>
		        <head>
		        <title>timeOut</title>
		        <LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		        </head>
		        <body topmargin="0" leftmargin="0" class="regular">
            <%
	        leftPos = 10
	        topPos = 20
	        bgthis = "#5582d2"
	        else
                if media = "export" then
	            leftPos = 20
	            topPos = 20
	            bgthis = "#5582d2"
                else
                leftPos = 90
	            topPos = 62
	            bgthis = "#5582d2"
                end if
	       
                
                
                
                
            '**** LOADBAR %>
            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
            <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:400px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	        <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	        <img src="../ill/outzource_logo_200.gif" />
	        <br />
	        Forventet loadtid: 5-20 sek.
	        </td><td align=right style="padding-right:40px;">
	        <img src="../inc/jquery/images/ajax-loader.gif" />
	
	        </td></tr></table>

	        </div>

            <script src="inc/feriekalender_jav.js"></script>

            <%response.flush %>

            <%if media <> "export" then %>
	
	        <!--#include file="inc/dato2_b.asp"-->
	
	
            <%end if %>
       
	

		
	        <%
            if media <> "export" then
                '**** MAIN MENU *****************

                call menu_2014()

            end if

	end if


        '***********************************************************
        '**************** PERIOD VARIABLES *************************
        '***********************************************************
        '** HAS TO BE CALLED AFTER  #include file="inc/dato2_b.asp"
        '***********************************************************

        lastweek = 0
	    firstjan = "1/"&strMrd&"/"&ysel
	    startDato = firstjan 
	    
	    stdatoThis = ysel&"/10/2"
	    stdato = day(stdatoThis)&"/"&month(stdatoThis)&"/"&year(stdatoThis)
	    
	    startDatoPerSQL = year(firstjan) &"/"& month(firstjan) & "/"& day(firstjan)
	    
	    select case per_interval 
        case 1
        slutDato = dateAdd("m", 1, startDato)
        case 3 
	    slutDato = dateAdd("m", 3, startDato)
	    case else
	    slutDato = dateAdd("m", 12, startDato)
	    end select 
	    
	    antalDays = dateDiff("d", startDato, slutDato, 2,2)
	    
	    slutDatoPer = dateAdd("m", -1, slutDato)
	    
	    
	    slutDato_md = datepart("m", slutDatoPer, 2, 2)
	    slutDato_yy = datepart("yyyy", slutDatoPer, 2, 2)
	    
	    select case slutDato_md
	    case 1,3,5,7,8,10,12
	    slutDato_dd = 31
	    case 2
	            if slutDato_yy = 2000 OR slutDato_yy = 2004 OR slutDato_yy = 2008 OR slutDato_yy = 2012 OR slutDato_yy = 2016 OR slutDato_yy = 2020 OR slutDato_yy = 2024 then
	            slutDato_dd = 29        
	            else
	            slutDato_dd = 28
	            end if
	    case else
	    slutDato_dd = 30
	    end select
	    
	    
	    slutDatoPerSQL = slutDato_yy &"/"& slutDato_md &"/"& slutDato_dd

        '** Ferieår **'
        if strMrd < 5 then
        startDatoFEOSQL = (ysel - 1) &"/5/1"
        slutDatoFEOSQL = ysel &"/4/30"
        else 
        startDatoFEOSQL = ysel &"/5/1"
        slutDatoFEOSQL = (ysel+1) &"/4/30"
        end if



    '********************************************************************************************************
    '**************************  PAGE CONTENT ***************************************************************
	'********************************************************************************************************
                
    '** MAIN PAGE DIV 
    %>
    <div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">
	<% 
    
       

        


    '***************************************************
    '********** SEARCH FILTER **************************
    '***************************************************
    if media <> "export" then
    
    
    

    '*** OK FOR ADMIN LEVEL 2 og 6 også, da de allerede er tjekkt om de er temaledere.
	if level <= 2 OR level = 6 then
	oskrift = "Ferie, Afspad. & Sygdom's kalender"
	else
	oskrift = "Ferie & Afspad. kalender"
	end if
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	
    if media <> "print" then

	'call filterheader(0,0,930,pTxt)
    call filterheader_2013(40,0,1050,oskrift) %>
	<table border=0 cellspacing=0 cellpadding=10 width=100%>
	<form action="feriekalender.asp?menu=job" method="post" name="periode" id="periode">
    <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
    <tr>
    
    <%call progrpmedarb %>
    
    
    
       
	<td valign=top style="width:400px;">
    <b>Periode:</b><br /><br />
	<b>Fra:</b> 
	<select name="FM_start_mrd">>
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
	<b>År:</b>
        <select id="Select2" name="yuse">
        <%
	'ysel = now
	for y = -5 to 10 %>
    <%yShow = dateAdd("yyyy", y, now) 
        
        if cint(year(yShow)) = cint(ysel) then
        ySele = "SELECTED"
        else
        ySele = ""
        end if
    
    %>
	<option value="<%=datePart("yyyy", yShow, 2,2)%>" <%=ySele %>><%=datePart("yyyy", yShow, 2,2)%></option> 
	<%next %>
            
        </select>
      &nbsp;
      
    
      <select id="per_interval" name="per_interval" onchange="submit();">
        <option value="1" <%=per_i_enCHK %>>1 md. </option>
        <option value="3" <%=per_i_treCHK %>>3 md. </option>
        <option value="12" <%=per_i_tolvCHK %>>12 md.</option>
      </select> 
      

	</td>
	</tr>
        <track /><td colspan="4" align="right">  <input id="Submit1" type="submit" value=" Vis >> " /></td>
	
	
	</form></table>
    
    
    


	</td></tr></table>
	</div>
	<%end if 'media 
    '***** SEARCH FILTER END ******************************
        
        
        
        
        
        
    'YEAR SELECTED%>

    <br /><br /><br />

	<h4>Valgt år: <%=ysel%></h4>
	
	
	
	
	
	
	<% 
    '********************************************************************************************************
    '*********** PERIOD HEADER -- to be repeated 
    '********************************************************************************************************
	
	tTop = 0
	tLeft = 0

    select case per_interval
    case 1
	tWdth = 1250
	case 3
    tWdth = 2025
    case else
    tWdth = 1945 '1184
    end select
	
	'call tableDiv(tTop,tLeft,tWdth)
	
	%>
            <div style="position:relative; width:<%=tWdth%>px; background-color:#FFFFFF; padding:3px; top:0px; left:0px;">
	
	
	
	<table cellspacing=0 cellpadding=0 border=0 bgcolor=#8caae6>
	<tr bgcolor="#3B5998">
	    <td style="height:15px; width:105px; padding-left:3px; border-right:1px #CCCCCC solid;" class=alt>
            Uge<br />
            <img src="ill/blank.gif" width=105 height=1 border=0 /><br /></td>
	
	<%
	    
	    
	    '*** WEEKS ****
	   
	    select case per_interval
        case 1
        stTop = 39 '39
        leftvaltot = 540
        case 3  
        stTop = 39' 39
        leftvaltot = 1280
        case else
        sttop = 23 '23
        leftvaltot = 1220
        end select

      
        if media = "print" then
        sttop = sttop 
        end if
	  
	    'dim monthwidth
	    'redim monthwidth(12)
	    m = -1
        lastweek = 0
        wwdt_nul = 0
        wwdt_53 = 0
	    for d = 0 to antalDays
	    
	    if d = 0 OR d > daysInInterval then
	    newDate = startDato
	    weekd = weekday(startDato, 3)
	        
            if d = 0 then    
            wwdt = (7 - weekd) * 3
            wwdt_nul = wwdt
            else
            wwdt = (weekd * 3)
            end if
	    
        else
	    newDate = dateadd("d", 1, newDate)
	    wwdt = 21 '(2px + 1px border)
	    end if
	    
	    weeknum = datepart("ww", newDate,2,2)

        select case right(weeknum, 1)
        case 0,2,4,6,8
        bgcolor = "#3B5998"
        case else
        bgcolor = "#5582d2"
        end select

        if cint(weeknum) = 53 then
        weekd = weekday(newDate, 2)
	    wwdt_53 = (7 - weekd) * 3
        'wwdt_53
        wwdt = wwdt_53
        bdright = 1
        else
        bdright = 0
        end if 

        'if d > 79 then
        'select case per_interval
        'case 1
        'wwdt = wwdt
        'case 3
        'wwdt = 15
        'case else
        'wwdt = wwdt
        'end select
        'end if

	    if weeknum <> lastweek then
	    %>
	    <td class=alt_lille align=center colspan="<%=weekspan %>" style="width:<%=wwdt*faktor %>px; background-color:<%=bgcolor%>; border-right:<%=bdright%> #999999 solid;">
         <%if d <> 0 AND d < daysInInterval then %>
         
         <% 
       
        if weeknum <> 53 then
        %>
        <%=weeknum%>
        <%
	    end if
        %>

       

         <%end if %>
	    <img src="ill/blank.gif" width=<%=wwdt*faktor %> height=1 border=0 /><br /></td>
	    <%
	    
	  
	    lastweek = weeknum
	    end if
	    
	    next
	
	
	%>
	</tr>
    
	</table>
	

    
    
	
	
	<table cellspacing=0 cellpadding=0 border=0 bgcolor=#ffffff>
	<tr bgcolor="#EFf3ff">
	
	<td style="width:105px; height:25px; padding-left:3px; border-right:1px #cccccc solid; border-bottom:1px #cccccc solid;">Medarb./ Måned
	<img src='../ill/blank.gif' width="105" height="1" alt="" border="0"><br></td>
	
	<%
	    
	    'lastmth = 0
	    'firstjan = "1/1/"&ysel
	    dim monthwidth
	    redim monthwidth(12) 
	    for d = 0 to per_interval - 1 'antalDays 'midt december / midt i 3 md.
	    
	    if d = 0 then
	    newDate = startdato 'firstjan
	    else
	    newDate = dateadd("m", 1, newDate)
	    end if
	    
	    monthnum = datepart("m", newDate,2,2)
	    
	    if monthnum <> lastmonth OR d = 0 then
	        
	        select case cint(monthnum)
	        case 3,5,7,10
	        mwdt = 93
            bgthis = "#FFFFFF"
            monthwidth(d) = 31
            case 1,8
	        mwdt = 93
            bgthis = "#d6dff5"
            monthwidth(d) = 31
            case 12
            bgthis = "#FFFFFF"
            mwdt = 93 + wwdt_53
            monthwidth(d) = 31
	        case 2
	            if ysel = 2000 OR ysel = 2004 OR ysel = 2008 OR ysel = 2012 OR ysel = 2016 OR ysel = 2020 OR ysel = 2024 _
                OR ysel = 2028 OR ysel = 2032 OR ysel = 2036 OR ysel = 2040 then
	            mwdt = 87
                monthwidth(d) = 29
	            else 
	            mwdt = 84
                monthwidth(d) = 28
	            end if
            bgthis = "#EFf3ff"
	        case else
	        bgthis = "#EFf3ff"
            mwdt = 90
            monthwidth(d) = 30
	        end select
	        
            select case per_interval 
            case 1
            monthwidth(d) = monthwidth(d)
            case 3 
            monthwidth(d) = monthwidth(d)
            case else
            monthwidth(d) = 1
            end select

	        %>
	    <td align="center" colspan="<%=monthwidth(d) %>" style="width:<%=mwdt%>px; background-color:<%=bgthis%>; border-bottom:1px #cccccc solid;"><%=left(monthname(monthnum), 3)%> - <%=year(newdate)%>
	    <img src="ill/blank.gif" width="<%=mwdt*faktor %>" height="1" /></td>
	    <%
	    lastmonth = monthnum
	    end if
	    
	    next
	    
        %>
        </tr>
        <%




    '**** WEEKDAYS **** 
    if per_interval = 3 OR per_interval = 1 then 
     %>
     <tr><td bgcolor="#Eff3FF" style="border-right:1px #cccccc solid; border-bottom:1px #cccccc solid;">&nbsp;</td>
     <%
     for d = 0 to antalDays - 1
      
     
        if d = 0 OR d > daysInInterval then
	    newDate = startDato
	    
         
        else
	    newDate = dateadd("d", 1, newDate)
	    end if
        
      select case weekday(newDate, 2)
      case 1
      bgcolor = "#8caae6"
      case 2
      bgcolor = "#eff3ff"
      case 3
      bgcolor = "#d6dff5"
       case 4
      bgcolor = "#8caae6"
       case 5
      bgcolor = "#eff3ff"
       case 6
      bgcolor = "CCCCCC"
      case 7
      bgcolor = "#CCCCCC"
      end select  

     thisDate = datepart("d", newDate, 2,2)
     select case right(thisDate, 2)
     case 01,03,07,09
     show = 1
     case else
     show = 0
     end select

    

     %>
      <td align="center" style="width:<%=2*faktor %>px; background-color:<%=bgcolor%>; border-bottom:1px #cccccc solid; font-size:8px;">
      <img src="ill/blank.gif" width="<%=2*faktor %>" height="1" alt="" border="0" /><br />
       <!--left(weekdayname(weekday(newDate, 2)), 1) -->
        
       <%=datepart("d", newDate, 2,2) %>
     
      </td>
      <%

     next%>
     </tr>
     <%end if 
	
     
    

     end if 'export







    '********************************************************************************************************
    '**** MAIN ROW SECTION FOR EACH EMPLOYEE *********************
    '********************************************************************************************************


	
	
	        dim medarbid, normTimerUge, medarbNavn, medarbNr, intFerieSaldo, intFerieFridageOpt, intFerieFridageAU, intFerieFridageSaldo
            dim intMidsFePl, intMidsFeAf, intMidsSyg, intMidsBarnSyg, intMidsAfspad, intMidsFeFri, intMidsFeAfUL, intFerieOpt, intFerieAU
            dim intMidsFeUb, intMidsFeFriUb, intMidsRejsedage
            dim dageVal_14, dageVal_13, dageVal_16, dageVal_17, dageVal_19 
	        redim medarbid(700), normTimerUge(700), medarbnavn(700), medarbNr(700), medarbInit(700)
            redim intMidsFePl(700), intMidsFeAf(700), intMidsSyg(700), intMidsBarnSyg(700), intMidsAfspad(700)
            redim intMidsFeFri(700), intMidsFeAfUL(700), intFerieOpt(700), intFerieSaldo(700), intFerieAU(700)
            redim intMidsFeUb(700), intMidsFeFriUb(700)
            redim intFerieFridageOpt(700), intFerieFridageAU(700), intFerieFridageSaldo(700)
            redim dageVal_14(700,12), dageVal_13(700,12), dageVal_16(700,12), dageVal_17(700,12), dageVal_19(700,12), intMidsRejsedage(700) 
	
	        'm = 1
	
	        for m = 0 TO UBOUND(intMids)
	
	        if intMids(m) <> 0 then
	        strSQLmed = "SELECT mid, mnavn, mnr, init FROM medarbejdere WHERE mid = "& intMids(m) 
	
	        oRec.open strSQLmed, oConn, 3
	        if not oRec.EOF then
	
	        'medarbid(m) = oRec("mid")
	         medarbnavn(m) = oRec("mnavn") 
             medarbNr(m) = oRec("mnr")
             medarbInit(m) = oRec("init")
             'medarbType(m) = oRec("medarbejdertype")
	
	        if media <> "export" then
	        %>
	        <tr>
	            <td bgcolor="#ffffff" class=lille style="width:105px; height:40px; border-bottom:1px #cccccc solid; border-right:1px #cccccc solid;">
	             <%=left(oRec("mnavn"), 12) %> (<%=oRec("mnr") %>) - <%=oRec("init") %>
	            <img src="ill/blank.gif" width="105" height="1" /><br /></td>
	    
	            <%
               

	            for d = 0 to per_interval - 1
	        
	                if lasttdbg = "#ffffff" then
	                bgtd = "#d6dff5"
	                else
	                bgtd = "#ffffff"
	                end if
	        
	                lasttdbg = bgtd
	            %>
	            <td bgcolor="<%=bgtd %>" colspan="<%=monthwidth(d) %>" style="border-bottom:1px #cccccc solid; border-right:0px #3B5998 solid;">
                <img src='../ill/blank.gif' width='1' height='1' alt='' border='0'>
                    </td>
	            <%

                
	            next
	            
             %>
	
	        </tr>
	        <%
             end if ' media
	        
	        end if
	        oRec.close
	
	end if
	
	next

     if media <> "export" then
	 %>
	</table>
	
	<% end if

    '**** END EMPLOYEES











    '********************************************************************************************************************************
    'MAIN LOOP FOR DOTS
    
    '*******************************************************************************************************************************

	
    if vis = 1 then '** Ferie rapport 

    sqlTypKri = "tfaktim = 14 OR tfaktim = 16 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 19"

    else
	
	if level <= 2 OR level = 6 then
	sqlTypKri = "tfaktim = 11 OR tfaktim = 14 OR tfaktim = 16 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 31 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 18 OR tfaktim = 19 OR tfaktim = 125"
	else
	sqlTypKri = "tfaktim = 11 OR tfaktim = 14 OR tfaktim = 16 OR tfaktim = 31 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 18 OR tfaktim = 19 OR tfaktim = 125"
	end if
	
    if lto = "cst" then
    sqlTypKri = sqlTypKri & " OR tfaktim = 7"
    end if
    
    end if
     
	
	'** Henter ferie og sygdom **' 
	i = 0
	tpp = 1
    feaf = 0

    select case vis 'Gruop total sum på export
    case 1
    sqlGrpBy = "tdato, tfaktim, tmnr"
    tdatofield = ""
    case 2
    sqlGrpBy = "tdato, tfaktim, tmnr"
    tdatofield = ""
    case else
    sqlGrpBy = "tdato, tfaktim, tmnr"
    tdatofield = ""
    end select

    lastMid = 0
   
	for m = 0 to UBOUND(intMids)
	    
	    
	    if intMids(m) <> 0 then


        intFerieOpt(m) = 0
        '** Ferie optjent i ferieår incl. ferie optj. u løn og overført ***'
        strSQLfeo = "SELECT sum(timer) AS timer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & intMids(m)  & ""_
	    &" AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 112) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' GROUP BY tmnr ORDER BY tdato"
        
        oRec.open strSQLfeo, oConn, 3
	    if not oRec.EOF then
        
       

         call normtimerPer(intMids(m) , oRec("tdato"), 6, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer/antalDageMtimer
         else
         ntimPerUse = 1
         end if 

          intFerieOpt(m) = oRec("timer") / ntimPerUse

        end if
        oRec.close
        
        
     


        intFerieAU(m) = 0
        '** Ferie afholdt + Uløn og udbetalt i ferieår ***'

        'if lto = "cst" then
        strSQLfeau = "SELECT timer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & intMids(m)  & ""_
	    &" AND (tfaktim = 14 OR tfaktim = 19 OR tfaktim = 16) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' ORDER BY tdato"
        'else
        'strSQLfeau = "SELECT timer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & intMids(m)  & ""_
	    '&" AND (tfaktim = 14 OR tfaktim = 16) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' ORDER BY tdato"
        'end if
	    

        'if session("mid") = 1 then
        'Response.Write strSQLfeau
        'Response.flush
        'end if

        oRec.open strSQLfeau, oConn, 3
	    while not oRec.EOF
        
       

         call normtimerPer(intMids(m) , oRec("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 


      
        intFerieAU(m) = intFerieAU(m) + oRec("timer") / ntimPerUse
       
          

          
        oRec.movenext
        wend
        oRec.close          


        '*** Saldo ***'
        intFerieSaldo(m) = intFerieOpt(m)/1 - (intFerieAU(m)/1)
       




        intFerieFridageOpt(m) = 0
        '** FerieFridage optjent i ferieår ***'
        strSQLfeo = "SELECT sum(timer) AS timer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & intMids(m)  & ""_
	    &" AND (tfaktim = 12) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' GROUP BY "& sqlGrpBy &" ORDER BY tdato"
	    
        'Response.Write strSQLfeo
        'Response.flush
        
        oRec.open strSQLfeo, oConn, 3
	    if not oRec.EOF then
        
       

         call normtimerPer(intMids(m) , oRec("tdato"), 6, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer/antalDageMtimer
         else
         ntimPerUse = 1
         end if 

          intFerieFridageOpt(m) = oRec("timer") / ntimPerUse

        end if
        oRec.close
        
        
        intFerieFridageAU(m) = 0
        '** FerieFridage afholdt og udbetalt i ferieår ***'
        strSQLfeau = "SELECT timer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & intMids(m)  & ""_
	    &" AND (tfaktim = 13 OR tfaktim = 17) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' ORDER BY tdato"
	    
        'Response.Write strSQLfeau
        'Response.flush
        
        oRec.open strSQLfeau, oConn, 3
	    while not oRec.EOF
        
       

         call normtimerPer(intMids(m) , oRec("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

          intFerieFridageAU(m) = intFerieFridageAU(m) + oRec("timer") / ntimPerUse

          
        oRec.movenext
        wend
        oRec.close          


        '*** Saldo ***'
        intFerieFridageSaldo(m) = intFerieFridageOpt(m)/1 - (intFerieFridageAU(m)/1)




	    
        lastMth = ""

        '***** MAIN SQL ***
        strSQLt = "SELECT sum(timer) AS timer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & intMids(m)  & ""_
	    &" AND ("& sqlTypKri &") AND tdato BETWEEN '"& startDatoPerSQL &"' AND '"& slutDatoPerSQL &"' GROUP BY "& sqlGrpBy &" ORDER BY tdato"
	    
	    'Response.Write "antaldays" & antaldays & "<br>"
	    'Response.Write strSQLt & "<br>"
	    'Response.flush

        'if session("mid") = 1 then
        'Response.Write strSQLt
        'Response.flush
        'end if

	    
	    oRec.open strSQLt, oConn, 3
	    while not oRec.EOF 

               
	        
	            '*** Normeret uge til (til bl.a ferie beregning + gns. fuld dag) ***'
	            '** Standard normtimer / dag uanset helligdage. 7,4 (37 tim.) skal altid give 1 hel dag.
	           call normtimerPer(intMids(m), oRec("tdato"), 0, 0)
               nomrTimerPrDag = ntimper

               'Response.Write oRec("tdato") & "ntim: "& ntimper &"  erH:"& erHellig & "<br>"

               if len(trim(nomrTimerPrDag)) <> 0 AND nomrTimerPrDag <> 0 then
               nomrTimerPrDag = nomrTimerPrDag
               nomrTimerPrDagTxt = nomrTimerPrDag
               else
               nomrTimerPrDag = 1
               nomrTimerPrDagTxt = 0
               end if
               

               mtThis = month(oRec("tdato"))
	     
	     select case oRec("tfaktim")
	     
	     case 11,18 'Ferie planlagt, Feriefridag pl
	     bgcol = "silver"
         ill = "dot_graae.gif"
	     tpekstra = -20
	     zinx = 300
	     'tnavn = "Ferie/Feriefrid. planlagt"
         call akttyper(oRec("tfaktim"), 1)
	     tnavn = akttypenavn '"Feriefridage brugt. + udbetalt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer  ~ "& formatnumber(oRec("timer")/(nomrTimerPrDag), 2) & " dag."
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsFePl(m) = intMidsFePl(m) + dageVal/1 
         
        
          
	     case 14 'Ferie afholdt 
	     bgcol = "green"
         ill = "dot_gron.gif"
	     tpekstra = -20
	     zinx = 400
	     'tnavn = "Ferie afholdt"
         call akttyper(oRec("tfaktim"), 1)
	     tnavn = akttypenavn '"Feriedage brugt. + udbetalt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer ~ "& formatnumber(oRec("timer")/(nomrTimerPrDag), 2) & " dag. ("& nomrTimerPrDag &")"
	     feaf = feaf + 1
         
         if vis = 1 then
         dageVal_14(m, mtThis) = dageVal_14(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         intMidsFeAf(m) = intMidsFeAf(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         else
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsFeAf(m) = intMidsFeAf(m) + dageVal/1
         end if

         

         case 16 'Ferie Udbetalt
         bgcol = "pink"
         ill = "dot_darkpink.gif"
	     tpekstra = -20
	     zinx = 400
	     call akttyper(oRec("tfaktim"), 1)
	     tnavn = akttypenavn '"Ferie udbetalt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer ~ "& formatnumber(oRec("timer")/(nomrTimerPrDag), 2) & " dag. ("& nomrTimerPrDag &")"
	     feaf = feaf + 1
         
         if vis = 1 then
         dageVal_16(m, mtThis) = dageVal_16(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         intMidsFeUb(m) = intMidsFeUb(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         else
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsFeUb(m) = intMidsFeUb(m) + dageVal/1
         end if

         
       
	     
         case 19 'Ferie u løn
	     bgcol = "yellowgreen"
         ill = "dot_yellowgron.gif"
	     tpekstra = -20
	     zinx = 400
	     tnavn = "Ferie afholdt u. løn"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer ~ "& formatnumber(oRec("timer")/(nomrTimerPrDag), 2) & " dag. ("& nomrTimerPrDag &")"
	     feaf = feaf + 1
	     
         if vis = 1 then
         dageVal_19(m, mtThis) = dageVal_19(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         intMidsFeAfUL(m) = intMidsFeAfUL(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         else
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsFeAfUL(m) = intMidsFeAfUL(m) + dageVal/1
         end if

         

	     case 20 'Syg
	     bgcol = "red"
          ill = "dot_rod.gif"
	     tpekstra = 0
	     zinx = 200
	     tnavn = "Syg"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
         dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsSyg(m) = intMidsSyg(m) + dageVal/1
	     
	     case 21 'Barn syg
	     bgcol = "orange"
          ill = "dot_orange.gif"
	     tpekstra = 10
	     zinx = 100
	     tnavn = "Barn syg"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
	     intMidsBarnSyg(m) = intMidsBarnSyg(m) + dageVal/1
	     
	     case 31 'Afsp. 
	     bgcol = "#8cAAe6"
          ill = "dot_blaa.gif"
	     tpekstra = -10
	     zinx = 100
	     tnavn = "Afspadsering"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsAfspad(m) = intMidsAfspad(m) + dageVal/1

	     case 13'Feriefri brugt
	     bgcol = "yellow"
         ill = "dot_gul.gif"
	     tpekstra = -10
	     zinx = 200
         call akttyper(oRec("tfaktim"), 1)
	     tnavn = akttypenavn '"Feriefridage brugt. + udbetalt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
         dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)

         if vis = 1 then
         dageVal_13(m, mtThis) = dageVal_13(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         intMidsFeFri(m) = intMidsFeFri(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         else
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsFeFri(m) = intMidsFeFri(m) + dageVal/1
         end if

         


         case 17 'Feriefri Udbetalt
	     bgcol = "#D592E1"
         ill = "dot_lightpink.gif"
	     tpekstra = -10
	     zinx = 200
         call akttyper(oRec("tfaktim"), 1)
	     tnavn = akttypenavn '"Feriefridage brugt. + udbetalt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
         dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)

         if vis = 1 then
         dageVal_17(m, mtThis) = dageVal_17(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         intMidsFeFriUb(m) = intMidsFeFriUb(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
         else
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsFeFriUb(m) = intMidsFeFriUb(m) + dageVal/1
         end if

         

	    


         case 125 'Rejsedage 
	     bgcol = "#E7A1EF"
         ill = "dot_E7A1EF.gif"
	     tpekstra = -20
	     zinx = 400
	     
         call akttyper(oRec("tfaktim"), 1)
	     tnavn = akttypenavn '"Rejsedage"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer ~ "& formatnumber(oRec("timer")/(nomrTimerPrDag), 2) & " dag. ("& nomrTimerPrDag &")"
	     feaf = feaf + 1
         
	     dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
         intMidsRejsedage(m) = intMidsRejsedage(m) + dageVal/1
       

         end select



         if lto = "intranet - local" OR lto = "cst" then
            if oRec("tfaktim") = 7 then

             bgcol = "#E7A1EF"
             ill = "dot_E7A1EF.gif"
	         tpekstra = 0
	         zinx = 200
	         tnavn = "Flextid brugt."
	         hoursval = formatnumber(oRec("timer"), 2) & " timer."
             dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)


         
            end if
         end if 
	     
	    
        'if media <> "export" then 
        'select case lto
        'case "wwf"
        ''intFerieSaldo(m) = intFerieSaldo(m) - (intMidsFeAfUL(m))
        'intFerieSaldo(m) = intFerieOpt(m)/1 - ((intMidsFeAf(m)/1 + intMidsFeUb(m)/1) + (intMidsFeAfUL(m)/1))

        ''Response.Write "her: " & intFerieOpt(m)/1 - ((intMidsFeAf(m)/1 + intMidsFeUb(m)/1) + (intMidsFeAfUL(m)/1)) & "<br>"
        'end select

        'end if


	     '** Beregn placering ***'
	     select case datepart("m", oRec("tdato"), 2,2)
	     case 1
	     bregnM = 0
	     mthval = 0
	     case else 
	     bregnM = dateadd("m", -1, oRec("tdato"))
	     mthval = (cint(datepart("m", bregnM, 2,2)) * 88) + cint(datepart("m", bregnM, 2,2)) * 4
	     end select

         daysDiff = dateDiff("d", startDato, oRec("tdato"), 2, 2)

         yearDiff = dateDiff("yyyy", startDato, oRec("tdato"), 2, 2)

         weekThis = datepart("ww", oRec("tdato"),2,2)
	     
	     stval = 109
	     dayval = daysDiff * 3
	     

         wwdt_nul = 3
         if yearDiff <> 0 then
         leftval = stval + (wwdt_nul + wwdt_53 +  dayval * faktor)
         else
         leftval = stval + (wwdt_nul + dayval * faktor)
         end if
	     

         if vis <> 2 then '** 2 = sum total
             select case vis
             case 1 

             if LastMid <> intMids(m) then
             'ekspTxt = ekspTxt &";xx99123sy#z"
                
                'if m <> 0 then
                'ekspTxt = ekspTxt &";;;;;;;;"& intFerieSaldo(LastMid)
                'end if
             if m <> 0 AND LastMid <> 0 then

             ekspTxt = ekspTxt & ekspTxtDt & ekspTxt14 & ekspTxt16 & ekspTxt19 & ekspTxt13 & ekspTxt17

             ekspTxt = ekspTxt &";xx99123sy#z"
             ekspTxt = ekspTxt &";xx99123sy#z"
             ekspTxt = ekspTxt &";xx99123sy#z"
             ekspTxt = ekspTxt &";xx99123sy#z"
             ekspTxt = ekspTxt &";xx99123sy#z"
             ekspTxt = ekspTxt &";xx99123sy#z"

             
             end if

             ekspTxt = ekspTxt & strOskrifter &";xx99123sy#z"
	         ekspTxt = ekspTxt & medarbnavn(m) &";"& medarbnr(m) &";"& medarbinit(m) &";"& formatnumber(intFerieOpt(m),2) &";;;;;"& formatnumber(intFerieSaldo(m), 2) &";;" & formatnumber(intFerieFridageOpt(m),2) & ";;;"& formatnumber(intFerieFridageSaldo(m), 2)
             LastMid = intMids(m)
                        
                        ekspTxtDt = ""
                        ekspTxt13 = ""
                        ekspTxt14 = ""
                        ekspTxt16 = ""
                        ekspTxt17 = ""
                        ekspTxt19 = ""

                        

             
             end if

             if lastMth <> monthname(oRec("month")) then

             
             ekspTxt = ekspTxt & ekspTxtDt & ekspTxt14 & ekspTxt16 & ekspTxt19 & ekspTxt13 & ekspTxt17

             ekspTxtDt = ";xx99123sy#z"
             ekspTxtDt = ekspTxtDt &";;;;"
             ekspTxtDt = ekspTxtDt & monthname(oRec("month")) &" "& year(oRec("tdato")) 
             
                        
                        ekspTxt13 = ";;;;"
                        ekspTxt14 = ";"
                        ekspTxt16 = ";"
                        ekspTxt17 = ";"
                        ekspTxt19 = ";"

                        'dageVal_14(m) = 0

             end if

             
             
                 select case oRec("tfaktim")
                 case 14
                 ekspTxt14 = ";"& dageVal_14(m, mtThis)
                 case 16
                 ekspTxt16 = ";"& dageVal_16(m, mtThis)
                 case 19
                 ekspTxt19 = ";"& dageVal_19(m, mtThis)
                 case 13
                 ekspTxt13 = ";;;;"& dageVal_13(m, mtThis)
                 case 17
                 ekspTxt17 = ";"& dageVal_17(m, mtThis)
                 end select
                 
                 'ekspTxt = ekspTxt & "("& oRec("tfaktim") &")"
                


             'case 2 
	         case else
             ekspTxt = ekspTxt & medarbnavn(m) &";"& medarbnr(m) &";"& medarbinit(m) &";"& oRec("tdato") &";"& tnavn &";"& oRec("timer") &";"& dageVal &";xx99123sy#z"
             end select
         end if
	     
         if media <> "export" then
        
        if media <> "print" then
        rowhgt = 41
        else
        rowhgt = 40
        end if



             %>
	   
	    <div class=hand id="divid_<%=i%>" style="position:absolute; left:<%=leftval%>px; top:<%=sttop + (tpp*rowhgt) + tpekstra%>px; background-color:#FFFFFF; height:10px; width:<%=2*faktor %>px;" onclick="showdatoinfo('<%=i%>')"><img src='../ill/<%=ill %>' width='<%=2*faktor %>' height='10' alt='<%=oRec("tdato") %>' border='0'></div>
	    
	    <div class=hand id="divid_info_<%=i%>" style="position:absolute; visibility:hidden; display:none; left:<%=(leftval)-50%>px; top:<%=25 + (tpp*rowhgt) + tpekstra - 20%>px; background-color:#ffffff; width:150px; border:2px yellowgreen solid; padding:5px; z-index:1000;" onclick="hidedatoinfo('<%=i%>')">
	    <%=medarbnavn(m) %><br />
        <b><%=ucase(left(weekdayname(datepart("w", oRec("tdato"))), 1)) &""& mid(weekdayname(datepart("w", oRec("tdato"))), 2,2) %>. d. <%=oRec("tdato") %></b><br />
	    Uge <%=datepart("ww", oRec("tdato"), 2,2) %> 
	    <br /><%=tnavn %><br />
	    <%=hoursval%> 
	   </div>
	    
	    <% end if
	    i = i + 1
	    
        lastMth = monthname(oRec("month"))
	    
	    oRec.movenext
	    wend
	    oRec.close
	    
        

        if len(trim(intMidsFePl(m))) <> 0 then
        intMidsFePl(m) = formatnumber(intMidsFePl(m))
        else
        intMidsFePl(m) = 0
        end if

        if len(trim(intMidsFeAf(m))) <> 0 then
        intMidsFeAf(m) = formatnumber(intMidsFeAf(m))
        else
        intMidsFeAf(m) = 0
        end if

        if len(trim(intMidsFeUb(m))) <> 0 then
        intMidsFeUb(m) = formatnumber(intMidsFeUb(m))
        else
        intMidsFeUb(m) = 0
        end if

        

        if len(trim(intMidsSyg(m))) <> 0 then
        intMidsSyg(m) = formatnumber(intMidsSyg(m))
        else
        intMidsSyg(m) = 0
        end if

        if len(trim(intMidsBarnSyg(m))) <> 0 then
        intMidsBarnSyg(m) = formatnumber(intMidsBarnSyg(m))
        else
        intMidsBarnSyg(m) = 0
        end if

        if len(trim(intMidsAfspad(m))) <> 0 then
        intMidsAfspad(m) = formatnumber(intMidsAfspad(m))
        else
        intMidsAfspad(m) = 0
        end if

        if len(trim(intMidsFeFri(m))) <> 0 then
        intMidsFeFri(m) = formatnumber(intMidsFeFri(m))
        else
        intMidsFeFri(m) = 0
        end if
        
         if len(trim(intMidsFeAfUL(m))) <> 0 then
        intMidsFeAfUL(m) = formatnumber(intMidsFeAfUL(m))
        else
        intMidsFeAfUL(m) = 0
        end if

        if len(trim(intFerieOpt(m))) <> 0 then
        intFerieOpt(m) = formatnumber(intFerieOpt(m))
        else
        intFerieOpt(m) = 0
        end if

        if len(trim(intMidsFeFriUb(m))) <> 0 then
         intMidsFeFriUb(m) = formatnumber(intMidsFeFriUb(m))
        else
         intMidsFeFriUb(m) = 0
        end if

         

         if len(trim(intFerieFridageOpt(m))) <> 0 then
        intFerieFridageOpt(m) = formatnumber(intFerieFridageOpt(m))
        else
        intFerieFridageOpt(m) = 0
        end if
        
        'if (intFerieOpt(m)) <> 0 then
        'intFerieSaldo(m) = intFerieOpt(m)/1 - (intMidsFeAf(m)/1 + intMidsFeAfUL(m)/1)
        'end if


       

        if len(trim(intFerieSaldo(m))) <> 0 then
        intFerieSaldo(m) = formatnumber(intFerieSaldo(m))
        else
        intFerieSaldo(m) = 0
        end if

        if len(trim(intFerieFridageSaldo(m))) <> 0 then
        intFerieFridageSaldo(m) = formatnumber(intFerieFridageSaldo(m))
        else
        intFerieFridageSaldo(m) = 0
        end if


         if len(trim(intMidsRejsedage(m))) <> 0 then
        intMidsRejsedage(m) = formatnumber(intMidsRejsedage(m))
        else
        intMidsRejsedage(m) = 0
        end if


            
        
        
        if media <> "export" then
        
        

        %>
        <div id="div1" style="position:absolute; left:<%=leftvaltot %>px; top:<%=(sttop-21)+((tpp)*41)%>px; width:650px; background-color:#FFFFFF; padding:2px; border-bottom:1px #999999 solid;">
        <table cellpadding=1 cellspacing=1 border=0 width=100%>

            <%if right(m, 1) = 0 then %>
       <tr>
           <td class="lille" style="white-space:nowrap;">Ferie optj.</td>
           <td class="lille" style="white-space:nowrap;">Afholdt <img src="../ill/dot_gron.gif" width="3" height="10" border="0" alt="Ferie Afholdt" /></td>
           <td class="lille" style="white-space:nowrap;">Udbetalt <img src="../ill/dot_darkpink.gif" width="3" height="10" border="0" alt="Ferie Udbetalt" /></td>
           <td class="lille" style="white-space:nowrap;">Afh. u. løn <img src="../ill/dot_yellowgron.gif" width="3" height="10" border="0" alt="Ferie Afholdt U. Løn" /></td>
           <td class="lille" style="white-space:nowrap;"><b>Fe. Saldo</b></td>
            <td class="lille" style="white-space:nowrap;">Planlagt <img src="../ill/dot_graae.gif" width="3" height="10" border="0" alt="Ferie planlagt og Feriefridage planlagt" /></td>
          
           <td class="lille" style="white-space:nowrap; padding-left:20px;">Feriefri</td>
           <td class="lille" style="white-space:nowrap;">Afh. <img src="../ill/dot_gul.gif" width="3" height="10" border="0" alt="Feriefridage afholdt" /></td>
           <td class="lille" style="white-space:nowrap;">Udb. <img src="../ill/dot_lightpink.gif" width="3" height="10" border="0" alt="Feriefridage udbetalt" /></td>
           <td class="lille" style="white-space:nowrap;"><b>FF. Saldo</b></td>
           <td class="lille" style="white-space:nowrap; padding-left:20px;">Syg <img src="../ill/dot_rod.gif" width="3" height="10" border="0" alt="Syg" /></td>
           <td class="lille" style="white-space:nowrap;">Barnsyg <img src="../ill/dot_orange.gif" width="3" height="10" border="0" alt="Barn syg" /></td>
           <td class="lille" style="white-space:nowrap;">Afspad. <img src="../ill/dot_blaa.gif" width="3" height="10" border="0" alt="Afspadsering" /></td>
           <td class="lille" style="white-space:nowrap;">Rejsedage. <img src="../ill/dot_E7A1EF.gif" width="3" height="10" border="0" alt="Rejsedage" /></td>

       </tr>
            <%else %>

            <tr>
            <td class="lille" style="white-space:nowrap;width:47px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:38px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:45px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:55px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:47px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:43px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:53px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:25px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:28px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:47px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:45px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:45px;">&nbsp;</td>
                 <td class="lille" style="white-space:nowrap;width:45px;">&nbsp;</td>
                <td class="lille" style="white-space:nowrap;width:45px;">&nbsp;</td>
       </tr>

            <%end if %>
        <tr>
            <td class="lille" align="right"><%=intFerieOpt(m) %></td>
            <td class="lille" align="right"><%=intMidsFeAf(m) %></td>
            <td class="lille" align="right"><%=intMidsFeUb(m) %></td>
            <td class="lille" align="right"><%=intMidsFeAfUL(m) %></td>
            <td class="lille" align="right"><b><%=intFerieSaldo(m) %></b></td>
            <td class="lille" align="right"><%=intMidsFePl(m) %></td>

            <td class="lille" align="right" style="padding-left:20px;"><%=intFerieFridageOpt(m) %></td>
            <td class="lille" align="right"><%=intMidsFeFri(m) %></td>
            <td class="lille" align="right"><%=intMidsFeFriUb(m) %></td>

            <td class="lille" align="right"><b><%=intFerieFridageSaldo(m) %></b></td>
            

            <td class="lille" align="right" style="padding-left:20px;"><%=intMidsSyg(m) %></td>
            <td class="lille" align="right"><%=intMidsBarnSyg(m) %></td>
            <td class="lille" align="right"><%=intMidsAfspad(m) %></td>
            <td class="lille" align="right"><%=intMidsRejsedage(m) %></td>
        </tr>


            

        </table>
        </div>
	    
        <%
        
        else

         if vis = 2 then '*** tot sum
            
             ekspTxt = ekspTxt & medarbnavn(m) &";"& medarbnr(m) &";"& medarbinit(m) &";"& intFerieOpt(m) &";"& intMidsFeAf(m) &";"& intMidsFeAfUL(m) &";"& intFerieSaldo(m) &";"& intMidsFePl(m) &";"& intMidsSyg(m) &";"& intMidsBarnSyg(m) &";"& intMidsAfspad(m) &";"& intFerieFridageOpt(m) &";"& intMidsFeFri(m) &";"& intFerieFridageSaldo(m) &";xx99123sy#z"
             
        
         end if
       
        end if
	    
	    tpp = tpp + 1
	    end if
	    

     
	
	next


    '**** END LOOP FOR DOTS


	
	 if media <> "export" then
	%>
	</div><!-- table div-->
	
	<%
    end if








    '********************************************************************************************************
    '******************* Export FOR CSV SECTION **************************' 
    '********************************************************************************************************


    'Response.Write ekspTxt
    'Response.end

if media = "export" then

    
    if vis = 1 then
    'ekspTxt = ekspTxt &";xx99123sy#z"
    'ekspTxt = ekspTxt &";;;;;;;;"& intFerieSaldo(LastMid)
     ekspTxt = ekspTxt & ekspTxtDt & ekspTxt14 & ekspTxt16 & ekspTxt19 & ekspTxt13 & ekspTxt17
    end if
   
    
    call TimeOutVersion()
    

	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	datointerval = ysel &", "& strMrd & " +"& per_interval & " måneder"
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\feriekalender.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				
				
				'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
                if vis <> 1 then
				objF.WriteLine(strOskrifter & chr(013))
                end if
				objF.WriteLine(ekspTxt)
				objF.close
				
				%>
				
	            <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Din CSV. fil er klar >></a>
	            </td></tr>
	            </table>
	            
	          
	            
	            <%
                
                
                Response.end
	            'Response.redirect "../inc/log/data/"& file &""	
				



end if%>

	
 <%
    if media <> "print" then

ptop = 40
pleft = 1085
pwdt = 200


	
	

pnteksLnk = "func="&func&"&FM_progrp="&progrp&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&per_interval="&per_interval&"&FM_start_mrd="&strMrd&"&yuse="&ysel


call eksportogprint(ptop,pleft,pwdt)
%>

        
        
        
    
      <tr>
        <td align=center>
        <a href="#" onclick="Javascript:window.open('feriekalender.asp?media=export&<%=pnteksLnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
        </td><td><a href="#" onclick="Javascript:window.open('feriekalender.asp?media=export&<%=pnteksLnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport (Dag/Dag)</a>


             </td>
       </tr>

        <!-- OPDELING PÅ Måned/MEdarB I EKSPORT --> 
        <tr>
        <td align=center>
        <a href="#" onclick="Javascript:window.open('feriekalender.asp?media=export&<%=pnteksLnk%>&vis=1', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
        </td><td><a href="#" onclick="Javascript:window.open('feriekalender.asp?media=export&<%=pnteksLnk%>&vis=1', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport <br />(Ferie rapport Md/Md)</a></td>
       </tr>
       

         <tr>
        <td align=center>
        <a href="#" onclick="Javascript:window.open('feriekalender.asp?media=export&<%=pnteksLnk%>&vis=2', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
        </td><td><a href="#" onclick="Javascript:window.open('feriekalender.asp?media=export&<%=pnteksLnk%>&vis=2', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport (Sum)</a></td>
       </tr>

    <tr>
   <td align=center><a href="feriekalender.asp?media=print&<%=pnteksLnk%>" target="_blank"  class='rmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
    </td><td><a href="feriekalender.asp?media=print&<%=pnteksLnk%>" target="_blank" class="rmenu">Print version</a></td>
   </tr>
   
	
   </table>
</div>
	


   

      <%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if


       itop = 150
       ileft = -10 
       iwdt = 400
       call sideinfo(itop,ileft,iwdt)%>
	    
        *) Feriesaldo og Ferieoptjent er opgjort for hele ferieåret, uanset forbrug i valgt periode.<br /><br />

        Ferie plantlagt og ferie afholdt indtastes via timeregistrerings siden. Det kræver at der er oprettet en
        aktivitet af typerne "Ferie plantlagt" og "Ferie afholdt".
        
        <br /><br />
          Fra den 1. januar 2001 optjenes 2,08 dages ferie for hver måneds ansættelse i optjeningsåret, som er lig kalenderåret. 
        Dette gælder også medarbejdere på del-tid.<br />
        2,08 * 12 måneder = 25 dage ell. 5 arbejdsuger.<br /><br />
        Ferie timer angives som timer og omregnes til dage udfra gns. normeret timer pr. uge 
        (se medarbejdertyper ell. medarb. afstem.). Så hvis man er ansat 37 timer, 
        angiver man altså 7,4 timer for en hel dags ferie, og 37 timer for en uge. 
        <br /><br />
        <b>Signatur forklaring:</b><br />
        (Norm. t./d. = 1 dag) <br />
        
         <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:silver; width:20px;"><img src="../ill/dot_graae.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Ferie planlagt + Feriefridage planlagt<br />
	       <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:green; width:20px;"><img src="../ill/dot_gron.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Ferie afholdt<br />
	      <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:#C91AE8; width:20px;"><img src="../ill/dot_darkpink.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Ferie udbetalt<br />
	     
               <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:yellowgreen; width:20px;"><img src="../ill/dot_yellowgron.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Ferie afholdt u. løn <br />
	    

         <%if level <= 2 OR level = 6 then 'progrp = 10%>
         <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:red; width:20px;"><img src="../ill/dot_rod.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Syg<br />
	      <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:orange; width:20px;"><img src="../ill/dot_orange.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Barn Syg<br />
	      <%end if %>
	      
          <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:#8caae6; width:20px;"><img src="../ill/dot_blaa.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Afspadsering<br />
	      <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:yellow; width:20px;"><img src="../ill/dot_gul.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Feriefridage brugt<br />
	        <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:#D592E1; width:20px;"><img src="../ill/dot_lightpink.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Feriefridage udbetalt<br />
	     
              <% if lto = "intranet - local" OR lto = "adra" then %>
          <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:#E7A1EF; width:20px;"><img src="../ill/dot_E7A1EF.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Rejsedage<br />
	        <%end if %>

          <% if lto = "xintranet - local" OR lto = "cst" then %>
          <span style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:#E7A1EF; width:20px;"><img src="../ill/dot_E7A1EF.gif" width="20" height="15" border="0" /></span>&nbsp;&nbsp; Flekstid brugt<br />
	        <%end if %>
	     
	   
	     
	    <br /><br />
	    <br /><br />
            &nbsp;
        
        <!-- side info slut -->
        </td></tr></table>
			</div>
			
			
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
            &nbsp;
		
	</div>	<!-- side div-->
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
