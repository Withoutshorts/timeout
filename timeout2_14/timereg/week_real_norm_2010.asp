<%response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/week_real_norm_inc.asp"-->

 <!--#include file="../inc/regular/topmenu_inc.asp"-->



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
    media = request("media")

    if len(trim(request("FM_eksporter_direkte"))) <> 0 then
    func = "export"
    media = "export"
    end if

   
    
    
	id = request("id")
	thisfile = "week_norm_2010.asp"

    if len(trim(request("nomenu"))) <> 0 then
    nomenu = request("nomenu")
    else
    nomenu = 0
    end if
	
    if nomenu <> 0 then
    rdir = "popuptreg"
    else
    rdir = "stat"
    end if

	call erStempelurOn()

    
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	'sendemail = request("sendemail")
	
	select case func 
	case "-"
	
	case else
	
	
	
	dim intMids
	redim intMids(1050) '150
	
	if len(trim(request("FM_useSogKri"))) <> 0 then
	useSogKri = 1
	else

        if len(request("FM_medarb")) <> 0 then

        useSogKri = 0

        else
        
	    if request.cookies("tsa")("af_sogkr") <> "" then     
        useSogKri = request.cookies("tsa")("af_sogkr") 
	    else
        useSogKri = 0
        end if

        end if

    end if
	
	if cint(useSogKri) <> 0 then
	skCHK = "CHECKED"
	else
	skCHK = ""
	end if


    if len(trim(request("FM_useSogKriAfs"))) <> 0 then
	useSogKriAfs = 1
	else
	        

        if len(request("FM_medarb")) <> 0 then 'er  der søgrt

        useSogKriAfs = 0

        else

        if request.cookies("tsa")("af_sogkr_afs") <> "" then     
        useSogKriAfs = request.cookies("tsa")("af_sogkr_afs") 
	    else
        useSogKriAfs = 0
        end if

        end if

	end if
	
	if cint(useSogKriAfs) <> 0 then
	skCHKafs = "CHECKED"
	else
	skCHKafs = ""
	end if


    if len(trim(request("FM_useSogKriGk"))) <> 0 then
	useSogKriGk = 1
	else
	    
        
         if len(request("FM_medarb")) <> 0 then 'er  der søgrt

        useSogKriGk = 0

        else
            
        if request.cookies("tsa")("af_sogkr_gk") <> "" then     
        useSogKriGk = request.cookies("tsa")("af_sogkr_gk") 
	    else
        useSogKriGk = 0
        end if

        end if

	end if
	
	if cint(useSogKriGk) <> 0 then
	skCHKgk = "CHECKED"
	else
	skCHKgk = ""
	end if
	
    response.cookies("tsa")("af_sogkr") = useSogKri
    response.cookies("tsa")("af_sogkr_afs") = useSogKriAfs
    response.cookies("tsa")("af_sogkr_gk") = useSogKriGk

    response.cookies("tsa").expires = date + 30




	if len(trim(request("FM_moreorless"))) <> 0 then
	moreorless = request("FM_moreorless")
	else
	moreorless = 0
	end if
	
	
    if moreorless <> "0" then
	mlSEL0 = ""
	mlSEL1 = "SELECTED"
	else
	mlSEL0 = "SELECTED"
	mlSEL1 = ""
	end if
	 
	
	if len(trim(request("FM_timekri"))) <> 0 then
	    
	    call erDetInt(request("FM_timekri"))
	    if isInt > 0 then
	    timeKri = 0
	    else
	    timeKri = request("FM_timekri")
	    end if
	
	else
	timeKri = 0
	end if
	
	if len(trim(request("FM_saldokri"))) <> 0 then
	saldoKri = request("FM_saldokri")
	else
	saldoKri = 0
	end if
	
	slSEL0 = ""
	slSEL1 = ""
	slSEL2 = ""
	
	select case saldoKri
	case 1
	slSEL1 = "SELECTED"
	case 2
	slSEL2 = "SELECTED"
	case else
	slSEL0 = "SELECTED"
	end select
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	usemrn = medarbid

    level = session("rettigheder")

    if func = "export" OR func = "print" then
    media = func
    end if

    'Response.write "func: "& func & " media: "& media &"<br>"

	

	if len(request("FM_medarb")) <> 0 OR media = "export" then
	
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
	
    lastMidtjk = 0
	
    call smileyAfslutSettings()

	'Response.Write request("FM_medarb")
	'Response.end
	
	'Response.Write "muse: "& request("muse") &"year(now) = trim(request(yuse)"& year(now) &" = "& trim(request("yuse"))
	
	if len(request("muse")) <> 0 then
    mnow = request("muse")

        if mnow > 12 then 'kvartal
        
        select case mnow 
        case 13
        mSele13 = "SELECTED" 
        case 14
        mSele14 = "SELECTED"
        case 15
        mSele15 = "SELECTED"
        case 16
        mSele16 = "SELECTED"
        case 17
        mSele17 = "SELECTED"
        end select

        end if

    else
    mnow = month(now)
    end if

   


    
    '** DEC 2014??
    '** maks frem til DD
    '** ÅTD tidligere år til 31.12...
    '** Afsluttet måender skal med på hver linje.

    'UGP = udgangspunkt Altid d. 1 ==> Ændret: altid sidste dag i md
    'ddato = idag, eller sidste dag i måned ==> Altid 1. dag i md
	
    useMthATD = month(now)
	if len(trim(request("yuse"))) = 0 then 'Altid nuværende MD
    
    ugp = now
    'ddato = datepart("d", ugp) &"-"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
    ddato = "1-"& datepart("m", ugp) &"-"& datepart("yyyy", ugp)
    opgjortprdato = dateadd("d",-1, ugp)
    showigartxt = "("&tsa_txt_319&")"

       
    else
        
            if cint(year(now)) = cint(request("yuse")) then
            
                   if cint(mnow) < cint(month(now)) then
               
                       tjkMth = mnow
                       call ds_ugp
                
                    else

                        '** If kvartal / ÅTD ***'
                        if mnow > 12 then
            
                        call mnowgt12

                        call ds_ugp

                        else


                        tjkMth = mnow 'datepart("m", now, 2,2)
                        call ds_ugp

                        end if
        
                    end if

            
       

             else

                     '** If kvartal ***'
                    if mnow > 12 then
                    
                                if cint(year(now)) > cint(request("yuse")) then
                                useMthATD = 12
                                end if
                                
                                call mnowgt12     
                                call ds_ugp

                    else

                    tjkMth = mnow
                    call ds_ugp

                    end if
       
             end if


            
        showigartxt = ""
        'ddato = datepart("d", ugp) &"/"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
       
        if mnow <= 12 then

             ddato = "1-"& datepart("m", ugp) &"-"& datepart("yyyy", ugp)

       
        end if

        opgjortprdato = ugp

        ''** Maks opgjort frem til igår ****'
        'Response.Write "her"
        if month(now) = month(ddato) AND year(now) = year(ddato) OR mnow = 17 then
        opgjortprdato = dateadd("d",-1, now)
        ugp = dateadd("d",-1, now)
        showigartxt = "("&tsa_txt_319&")"
        end if
    
    end if
	
	
	
	
	if media <> "print" AND media <> "export" then
	
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
            
            <%if nomenu <> "1" then  
            
            leftPos = 90
	        topPos = 102
            
            %>
	       


            <%call menu_2014() %>

            <%else 
            
            leftPos = 20
	        topPos = 20
            
            %>
            <%end if %>
	
	<%else 
	
	leftPos = 20
	topPos = -30
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
    
      if func = "export" OR func = "print" then
    media = func
    end if
    
    end if



        if func = "export" then
        forvloadtid = 3
        
        if mnow > 12 then
        forvloadtidTrip = 1.7
        else
        forvloadtidTrip = 0.5
        end if

        for m = 0 to UBOUND(intMids)
        forvloadtid = forvloadtid + forvloadtidTrip/1 
        next

        forvloadtid = formatnumber(forvloadtid, 0)

        else
        forvloadtid = "4-15"
        end if

 
    %>
	

    <script src="inc/week_real_norm_jav.js"></script>


     <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:300px; width:300px; background-color:#ffffff; border:10px #CCCCCC solid; padding:10px; z-index:100000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" /><br />&nbsp;
  
	</td></tr>
    <tr><td colspan=2>  <div id="load_cdown">Forventet loadtid: <%=forvloadtid %> sek.</div></td></tr>
    </table>

	</div>


    
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">
	
       
	<%

    if len(trim(request("FM_md_week"))) <> 0 then

      useMDorWeek = request("FM_md_week")
    else
      useMDorWeek = SmiWeekOrMonth

    end if


    if cint(useMDorWeek) = 1 then
    opdelNormMDWeekChk0 = "" 
	opdelNormMDWeekChk1 = "CHECKED"
    else
    opdelNormMDWeekChk0 = "CHECKED" 
	opdelNormMDWeekChk1 = ""
    end if


    if media <> "export" then

	oimg = "ikon_medarbaf_48l.png"
	oleft = 0
    otop = 0
    owdt = 600
    
    if cint(SmiWeekOrMonth) = 0 then
        peridoeTxt = tsa_txt_005
    else
        peridoeTxt = tsa_txt_430
    end if

    select case lto
    case "tec", "esn"
        oskrift = "Luk "& peridoeTxt &"(er) for medarbejdere"
    case else
        oskrift = "Godkend "& peridoeTxt &"(er) for medarbejdere"
    end select

	
	
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	end if

    end if
	
	
	
	if media <> "print" AND media <> "export" then
	
	call filterheader_2013(0,0,1200,oskrift)%>
	<table border=0 cellspacing=0 cellpadding=0 width=100%>
	<form action="week_real_norm_2010.asp" method="post" name="periode" id="periode" target="">
    <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
     <input id="Hidden1" name="nomenu" value="<%=nomenu %>" type="hidden" />
    
    <tr>
    
    <%call progrpmedarb %>
    
    
       
	<td valign=top style="width:300px; padding-top:5px;"><br /><b>Periode</b><br /><br />År:
        <select id="Select2" name="yuse">
        <%
	ysel = now
	for y = 0 to 5 %>
    <%yShow = dateAdd("yyyy", -y, ysel) 
        
        if year(yShow) = year(ugp) then
        ySele = "SELECTED"
        else
        ySele = ""
        end if
    
    %>
	<option value="<%=datePart("yyyy", yShow, 2,2)%>" <%=ySele %>><%=datePart("yyyy", yShow, 2,2)%></option> 
	<%next %>
            
        </select>
        <br /><br />
        Måned:<br />
        <select id="Select3" name="muse" size=5 style="width:200px;">
        
            <%for m = 1 to 12 
                
                if m = month(ugp) then
                mSele = "SELECTED"
                else
                mSele = ""
                end if 
                
            %>       
            <option value="<%=m %>" <%=mSele %>><%=monthname(m) %></option>
            <%next %>

             <option value="13" <%=mSele13 %>>1. kvartal</option>
            <option value="14" <%=mSele14 %>>2. kvartal</option>
            <option value="15" <%=mSele15 %>>3. kvartal</option>
            <option value="16" <%=mSele16 %>>4. kvartal</option>
            <option value="17" <%=mSele17 %>>ÅTD</option>

        </select>
        <br /><br />
        <input type="checkbox" id="FM_eksporter_direkte" name="FM_eksporter_direkte" value="1" />Eksporter direkte til excel (hurtigere load) 
          <br /><br />
       

      
        <span id="sp_filterava" style="color:#5582d2;">[+] <b>Filter:</b> </span>(avanceret)<br />


   


        <input type="hidden" id="div_filterava_show" value="0" />
        <div id="div_filterava" style="display:none; visibility:hidden;">
        <table style="padding:10px;">

            <tr><td><br /><b>Opdel norm på</b><br />
        
        

        Hele måneder <input type="radio" name="FM_md_week" value="1" <%=opdelNormMDWeekChk1 %> /><br />Hele uger<input type="radio" name="FM_md_week" value="0" <%=opdelNormMDWeekChk0 %> />  <br /><br />&nbsp;</td></tr>

        <tr><td>
        <input id="Checkbox1" name="FM_useSogKri" value="1" type="checkbox" <%=skCHK %> />Vis kun uger/måneder med 
        <select id="Select4" name="FM_moreorless">
            
            <option value="0" <%=mlSEL0 %>>mere</option>
            <option value="1" <%=mlSEL1 %>>mindre</option>
        </select> end 
        <input id="Text1" type="text" name="FM_timekri" value="<%=timeKri %>" style="width:30px;" /> time(r)
                </td></tr><tr><td>
        på deres 
         <select id="Select5" name="FM_saldokri">
            <option value="0" <%=slSEL0%>>Fleks (realiseret / normeret)</option>
            <option value="1" <%=slSEL1%>>Fleks (komme/gå / normeret)</option>
            <option value="2" <%=slSEL2%>>Balance (realiseret / komme/gå)</option> 
            
        </select> pr. uge
        </td></tr>
            </tr><td>
	    
                <br />
          <input id="Checkbox2" name="FM_useSogKriAfs" value="1" type="checkbox" <%=skCHKafs %> />Vis kun ikke afsluttede uger/måneder<br />
          <input id="Checkbox3" name="FM_useSogKriGk" value="1" type="checkbox" <%=skCHKgk %> />Vis kun ikke godkendte uger/måneder<br />
        
       </td></tr>         
                </table>
        </div>
    </td>
	</tr>
    <tr>
        <td valign="bottom" align=right colspan=3>
	     <input id="Submit1" type="submit" value=" Vis afstemning >> " />
	</td>
	</tr>
	
	
	</form></table>
    
    
    
    <!-- filter header sLut -->
	</td></tr></table>
	</div>
    
    
    <%
    
    else
   
    end if 'media
    

    call akttyper2009(6)
	akttype_sel = "#-99#, " & aktiveTyper
	akttype_sel_len = len(akttype_sel)
	left_akttype_sel = left(akttype_sel, akttype_sel_len - 2)
	akttype_sel = left_akttype_sel
    
    
	public anormTimerTot, arealTimerTot, atotalTimerPer100, aafspadUdbBalTot, aAfspadBalTot, arealfTimerTot
    public afradragTimerTot, altimerKorFradTot, aafspTimerTot, aafspTimerBrTot, aafspTimerUdbTot 
    public normtime_lontime, balRealNormtimer, balRealLontimer, bgc, rejsedage_tot
    'public normtime_lontimeAkk, balRealNormtimerAkk, korrektionRealTot, ferieFriAfVal_md_tot,  ferieAfVal_md_tot
	
    public ferieAfVal_md, sygDage_barnSyg, ferieAfVal_md_tot, ferieAfulonVal_md_tot, sygDage_barnSyg_tot, ferieFriAfVal_md, ferieFriAfVal_md_tot, normtime_lontimeAkk, balRealNormtimerAkk, korrektionRealTot, divfritimer_tot, omsorg_tot 
    public sygeDage_tot, barnSyg_tot, barsel_tot, lageTimer_tot, tjenestefri_tot, aldersreduktionBrTot
    public flexTimer_tot, sTimer_tot, adhocTimer_tot
    
    
    if media <> "print" AND media <> "export" then    		

	
	tTop = 20
    tLeft = 0
	tWdth = 780 '1280
	globalWdt = tWdth 

	'call tableDiv(tTop,tLeft,tWdth)

    %>
     <div name="medarbafstem_aar" id="medarbafstem_aar" style="position:relative; left:<%=tLeft%>px; top:<%=tTop%>px; display:; visibility:visible; width:95%; background-color:#FFFFFF; padding:20px;">
   
    <%

      
	
	else
	
	'tTop = 20
	'tLeft = 20
	'tWdth = 800
	
	end if
	
	
    showextended = 0
	call stempelur_kolonne(lto, showextended)


    ddatoEndBeregn = dateAdd("m", 1, ddato)
    ddatoEndBeregn = dateAdd("d", -1, ddatoEndBeregn) 
    weekThis = datePart("ww", ddatoEndBeregn, 2,2)

    if mnow <= 12 then
    weekStMth = datePart("ww", "1-"& month(ddato)&"-"&year(ddato), 2,2)
    'weekDiff_WeekReal_norm = dateDiff("ww", "1-"& month(ddato)&"-"&year(ddato), ddato, 2, 2)
    weekDiff_WeekReal_norm = dateDiff("ww", "1-"& month(ddato)&"-"&year(ddato), ugp, 2, 2) 
    else

        select case mnow
        case 13
             '** Q1 
             startMd = "1-1-"&year(ddato)
             endMd = "31-3-"&year(ddato)
             weekStMth = datePart("ww", "1-1-"&year(ddato), 2,2)
             weekDiff_WeekReal_norm = dateDiff("ww", "1-1-"&year(ddato), "31-3-"&year(ugp), 2, 2) 

        case 14

              '** Q2 
             startMd = "1-4-"&year(ddato)
             endMd = "30-6-"&year(ddato)
             weekStMth = datePart("ww", "1-4-"&year(ddato), 2,2)
             weekDiff_WeekReal_norm = dateDiff("ww", "1-4-"&year(ddato), "30-6-"&year(ugp), 2, 2) 

        case 15
            
              '** Q3 
             startMd = "1-7-"&year(ddato)
             endMd = "31-10-"&year(ddato)
             weekStMth = datePart("ww", "1-7-"&year(ddato), 2,2)
             weekDiff_WeekReal_norm = dateDiff("ww", "1-7-"&year(ddato), "30-9-"&year(ugp), 2, 2) 

        case 16

             '** Q4 
             startMd = "1-10-"&year(ddato)
             endMd = "31-12-"&year(ddato)
             weekStMth = datePart("ww", "1-10-"&year(ddato), 2,2)
             weekDiff_WeekReal_norm = dateDiff("ww", "1-10-"&year(ddato), "31-12-"&year(ugp), 2, 2)    

        case 17

         '**ÅTD 
         startMd = "1-1-"&year(ddato)
         endMd = now
         weekStMth = datePart("ww", "1-1-"&year(ddato), 2,2)
         weekDiff_WeekReal_norm = dateDiff("ww", "1-1-"&year(ddato), ugp, 2, 2) 

        end select

    end if
   

    'response.write "ugp: "& ugp &  " ddato: "& ddato  &"mnow: "& mnow &" weekStMth: "& weekStMth &" weekDiff_WeekReal_norm: "& weekDiff_WeekReal_norm
	'response.flush
	

	if media <> "export" then%>
	<!--- Table start -->
	
	<table cellpadding=0 cellspacing=0 border=0 width="100%">
     <tr><td colspan="5" valign=top style="padding:0px;">





    <%if mnow <= 12 then %>
	<h4><%=monthname(month(ddato)) &" "& year(ddato) %><br />
    <%else %>
	<h4><%=monthname(month(startMd))%> - <%=monthname(month(endMd)) &" "& year(endMd) %><br />
    <%end if %>

    <span style="color:darkred; font-size:9px;">Opgjort pr. <b><%=formatdatetime(opgjortprdato, 1) %></b> <%=showigartxt %></span></h4>
	
    </td></tr></table>
    
    <table cellpadding=0 cellspacing=0 border=0 width="100%">

    <%end if %>
    
    <%
    
    	
	strEksportTxt = "Medarbejder;Mnr;Initialer;Projektgrupper/afd.;År;Måned;Uge;"& tsa_txt_173 & " (hh:mm); "& tsa_txt_173 & " (100 digit);"
	    

	
	
	
	if session("stempelur") <> 0 then
        
        strEksportTxt = strEksportTxt & "Komme/Gå (løntimer, hh:mm); Komme/Gå (100 digit);"

        if showkgtil = 1 then
	    strEksportTxt = strEksportTxt & "Tillæg +/-;Komme/Gå tid Sum;" 
        end if
	
      strEksportTxt = strEksportTxt & tsa_txt_284 & "+/- (Komme-Gå/Norm. hh:mm); +/- (Komme-Gå/Norm. 100 digit);"
      strEksportTxt = strEksportTxt & tsa_txt_284 & "+/- Akkumuleret (Komme-Gå/Norm. hh:mm);+/- Akkumuleret (Komme-Gå/Norm. 100 digit);"
    
    end if
        


        strEksportTxt = strEksportTxt  & tsa_txt_172 & ";"

    	if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then
	    strEksportTxt = strEksportTxt & "Heraf fakturerbare timer;"
	    
        if cint(showkgtil) = 1 then 
        strEksportTxt = strEksportTxt  &"Korrektion Realiseret;"
        end if

          strEksportTxt = strEksportTxt & tsa_txt_284 & " +/- (Real. / Norm.);"
            strEksportTxt = strEksportTxt & tsa_txt_284 & " +/- Akkumuleret (Real. / Norm.);"
        
        end if
	    
	
   
	if session("stempelur") <> 0 then
	
	      
	
	         if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then
	        strEksportTxt = strEksportTxt & tsa_txt_284 & " (Real. / Lønt.);"
	        end if
	end if
	

    '*** Afspandsering / Overarbejede ***'
	if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then

              if lto <> "fk" AND lto <> "kejd_pb" AND lto <> "adra" AND lto <> "cisu" then 
	             strEksportTxt = strEksportTxt & tsa_txt_283 &" "& tsa_txt_164 & "(enh.);"
              end if
                
                 strEksportTxt = strEksportTxt & "Afspads.;"

	                if lto <> "fk" AND lto <> "kejd_pb" AND lto <> "adra" AND lto <> "cisu" then
                        if lto <> "tec" AND lto <> "esn" then
                        strEksportTxt = strEksportTxt & "Udbetalt;Ønsk. Udbe.;"
                        end if
                        strEksportTxt = strEksportTxt & tsa_txt_283 &" "& tsa_txt_280 & ";"
                    end if
	  
	end if


              


                'TEC / ESN special  
                select case lto
                case "tec", "esn", "intranet - local"
                strEksportTxt = strEksportTxt & global_txt_172 &";"& global_txt_148 &";"& global_txt_179 &";"   
                end select



                select case lto
                case "esn", "intranet - local"
                strEksportTxt = strEksportTxt & global_txt_132 &";"& global_txt_157 &";"
                end select


    select case lto
    case "tec", "esn"
    ferieFriTxt = "Særligferie"
    case else 
    ferieFriTxt = "Feriefridage Afholdt"
    end select
	

    strEksportTxt = strEksportTxt & "Ferie Afholdt;Ferie Afholdt u. løn;"& ferieFriTxt &";"

         select case lto
                case "xintranet - local", "fk" 
                 strEksportTxt = strEksportTxt &"1 maj timer;"
                end select
                
               


                select case lto
                case "xintranet - local", "fk", "kejd_pb" 
                 strEksportTxt = strEksportTxt &"Omsorgsdage ~ dage;"
                end select
    
            

                    select case lto
                case "xintranet - local", "fk", "kejd_pb"
                 strEksportTxt = strEksportTxt &"Tjenestefri ~ timer;"
                end select

                select case lto
                case "xintranet - local", "fk", "kejd_pb" 
                 strEksportTxt = strEksportTxt &"Barsel/Orlov ~ dage;"
                end select
                

                 select case lto
                case "xxintranet - local", "fk" 
                 strEksportTxt = strEksportTxt &"Læge ~ timer;"
                end select

                
                'Rejsedage 
                select case lto
                case "intranet - local", "adra"
                strEksportTxt = strEksportTxt & global_txt_194 &";"
                end select 


               '*** KUN ADMIN, HVIS DU ER TEAMLEDER (temaledere kan kun se denne side for de medarb. de er teamledere for, hvis ikke level = 1. TeamlederKri er tjekket i toppen) 
               if session("rettigheder") <= 2 OR session("rettigheder") = 6 OR (session("mid") = usemrn) then
               strEksportTxt = strEksportTxt & "Syg ~ dage;"
               end if

               if session("rettigheder") <= 2 OR session("rettigheder") = 6 OR (session("mid") = usemrn) then
               strEksportTxt = strEksportTxt & "Barn syg ~ dage;"
               end if

               

                strEksportTxt = strEksportTxt & "Afsluttet;Godkendt, 0:Afventer, 1:Godkendt, 2:Afvist"
    
    
    '******************************************************************************
    '** MAIN LOOP MEDARBEJDERE ***'
    '******************************************************************************

          if func = "export" then
       
	     %><div style="position:absolute; left:40px; top:220px; width:800px; z-index:-1;">
         <b>Henter data for medarbejder:</b>
          <%
         end if


    lastMid = 0
    'normtime_lontimeAkkGT = 0
	for m = 0 to UBOUND(intMids)
	


	       
	   
	 if cint(intMids(m)) <> 0 then


           if media <> "export" then
	    
	   
	    
	                    'select case right(m,1)
	                    'case 0,2,4,6,8
	                     if m <> 0 then%>
	                    </td></tr>
	                     <%end if %>

	                    <tr><td valign=top style="padding:0px;">
	  
	    
                        <%
	   
	               


            end if

            	 'Response.write "strEksportTxt: "& intMids(0) &" : "& strEksportTxt & "<br>"

	meTxt = ""
	call meStamdata(intMids(m))
	

   
	if media <> "export" then
                            


    if media <> "print" then
    tpd = 2      
    else
    tpd = 1
    end if
                                                    
    %>

    <br /><br />
	<h4><%=meNavn & " ["& meInit &"]"%></h4>
    <table cellspacing=0 cellpadding=<%=tpd %> border=0 width="100%"><tr>
    <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Uge</b> - dato</td>
	
	    <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;"><b><%=tsa_txt_173%></b></td>

          <%if session("stempelur") <> 0 then %> 
       <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;"><b>Komme / Gå tid</b><br />(løntimer)</td>

	  <%if showkgtil = 1 then %>
	 <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;"><b>Tillæg +/-</b><br />Ferie, <br />sygdom, <br />korrektion<br /> etc. indberettet</td>
	 <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;"><b>= Sum</b><br />Komme/gå tid<br /> + tillæg</td>
     <%end if %>
	 
    
    <%if lto <> "kejd_pb" then %>
	 <td align=right valign=bottom class=lille style="border:1px #DCF5BD solid; border-bottom:1px silver solid;"><b><%=tsa_txt_284%> +/-</b><br />Komme/gå tid<br /> / Normtid</td>
     <td bgcolor="#DCF5BD" valign=bottom class=lille style="border:0px silver solid; border-bottom:1px silver solid; white-space:nowrap;"><b><%=tsa_txt_284%> +/-<br / >Akkumuleret</b><br />Komme/gå tid<br /> / Normtid</td>
	 <%end if %>
	 
	 <%end if %>



      <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;">
       <%select case lto 
       case "kejd_pb"
       %>
       <b>Timer</b><br /> indberettet<br /> på aktiviteter
       <%
       case else
       %>
       <b>Timer<br /> realiseret</b>
      <br />(indberettet på aktivit.)
       <%
       end select %>
      
      </td>
	  <%if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then %>
	 <td class=lille align=right valign=bottom style="border-bottom:1px silver solid;">(heraf<br />fak.bare)</td>
    <%end if %>


      <%if cint(showkgtil) = 1 AND (lto <> "cst" AND lto <> "tec" AND lto <> "esn") then %>
	    <td class=lille valign=bottom style="border-bottom:1px silver solid;"><b>Korrektion</b><br />(overført saldo fra<br /> senest afsluttet<br /> løn-periode)</td>
        <%end if %>
	

        <%if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then %>
	  <td align=right valign=bottom class=lille style=" border:1px pink solid; border-bottom:1px silver solid;"><b><%=tsa_txt_284%> +/-</b>

           <%select case lto
           case "kejd_pb"%>
          <br /> Indberettet <br />/ Normeret
            <%case else %>  
          <br />Realiseret<br /> / Normeret
	        <%end select %>


	  </td>
       <td valign=bottom class=lille bgcolor="pink" style="border:0px pink solid; border-bottom:1px silver solid; white-space:nowrap;"><b><%=tsa_txt_284%> +/- <br / >Akkumuleret</b>

            <%select case lto
           case "kejd_pb"%>
          <br /> Indberettet <br />/ Normeret
            <%case else %>  
          <br />Realiseret<br /> / Normeret
	        <%end select %>

       </td>
        <%end if %>

	        <%if session("stempelur") <> 0 AND (lto <> "kejd_pb" AND lto <> "cst" AND lto <> "tec" AND lto <> "esn") then %>
	        <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;"><b>Balance +/-</b><br />Realiseret<br />/ Komme/gå tid</td>
            <%end if %>

	  <%  if m = 0 then
          globalWdt = globalWdt + 200
          end if
        %>

	

	 
	               <!-- Afspad / Overarb --->
	                   <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>

                <%if lto <> "fk" AND lto <> "kejd_pb" AND lto <> "adra" AND lto <> "cisu" then %>
	             <td align=right valign=bottom class=lille style="border-bottom:1px silver solid;"><b><%=tsa_txt_283%><br /> <%=tsa_txt_164%></b><br />(enh.)</td>
                <%end if %>
	                      <td align=right valign=bottom style="border-bottom:1px silver solid; white-space:nowrap;" class=lille><b>Afspads.</b></td>
                           
                           <%if lto <> "fk" AND lto <> "kejd_pb" AND lto <> "adra" AND lto <> "cisu" then 
                               
                               if lto <> "tec" AND lto <> "esn" then%>
	                               <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Udbetalt</b></td>
	                               <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Ønsk. Udbe.</b></td>
	                            <%end if %>
                           <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b><%=tsa_txt_283 &" "& tsa_txt_280 %></b></td>
                           <%end if %>
	       
	 
	             <%
                     if m = 0 then
                     globalWdt = globalWdt + 100
                     end if
	             end if 




      


      'TEC / ESN special  
      select case lto
       case "esn", "intranet - local"
        %>

           <!--<td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_147 %></b></td>-->
           <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_132 %></b></td>
           <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_157 %></b></td>

        <%
       end select



       select case lto
       case "tec", "esn", "intranet - local"
                     %>
                     <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_172 %></b><br />
	                ~ dage</td>
                     <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_148 %></b><br />
	                ~ timer</td>
                     <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_179 %></b><br />
	                ~ timer</td>
                    <%


       end select




    

             %>

     <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Ferie afholdt</b><br />
	 ~ dage</td>

          <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Ferie afholdt u. løn</b><br />
	 ~ dage</td>


      

        <%

      
    select case lto
    case "tec", "esn"
    ferieFriTxt = "Særligferie"
    case else 
    ferieFriTxt = "Feriefridage afholdt"
    end select %>
    
      <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b><%=ferieFriTxt %></b><br />
	 ~ dage<br />
     <% 
    select case lto
    case "tec", "esn"
    
    case else 
    %>(Incl. udbetalt)<%
    end select %>
     


      </td>
        <%   if m = 0 then
             globalWdt = globalWdt + 100
        end if %>
         


          <%select case lto
            case "xintranet - local", "fk" %>
         <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>1 Maj timer</b></td>

         <%
              if m = 0 then
              globalWdt = globalWdt + 50
              end if
             end select
             

             'Rejsedage 
            select case lto
            case "intranet - local", "adra"
            %>
        	 <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b><%=global_txt_194 %></b></td>

            <%
              if m = 0 then
              globalWdt = globalWdt + 50
              end if
            end select 

             
             select case lto
            case "xintranet - local", "fk", "kejd_pb" %>
        <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Omsorgs<br />dage</b><br />
	 ~ dage<br />
     </td>
    <%      

         if m = 0 then
         globalWdt = globalWdt + 50
         end if
        end select 


          select case lto
            case "xintranet - local", "fk", "kejd_pb" %>
        <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b><%=global_txt_179 %></b><br />
	 ~ timer<br />
     </td>
    <%      

         if m = 0 then
         globalWdt = globalWdt + 50
         end if
        end select %>



         <%
          select case lto
            case "xintranet - local", "fk", "kejd_pb" 
             globalWdt = globalWdt + 50%>
        <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Barsel/Orlov</b><br />
	 ~ dage<br />
     </td>
   
    <%
         if m = 0 then
         globalWdt = globalWdt + 50
         end if
        end select %>


            <%
          select case lto
            case "xxintranet - local", "fk" 
             globalWdt = globalWdt + 50%>
       
        <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Læge</b><br />
	 ~ timer<br />
     </td>
    <%
    
        end select %>



    	 
	  <%if level <= 2 OR level = 6 OR (session("mid") = usemrn) then %>
	 <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Syg</b><br />~ dage</td>
         <td valign=bottom style="border-bottom:1px silver solid;" class=lille><b>Barn syg</b><br />~ dage</td>
	 <%
                 if m = 0 then
                 globalWdt = globalWdt + 100
                 end if
         end if 

         
         
         select case lto
         case "tec", "esn"
         gkTxt = "lukket"
         case else
         gkTxt = "godkendt"
         end select%>
	 


	    <td style="border-bottom:1px silver solid;" valign=bottom class=lille><b><%=peridoeTxt %> <br />afsluttet?</b></td>
        <td style="border-bottom:1px silver solid;" valign=bottom class=lille><b><%=peridoeTxt %><br /><%=gkTxt %>?</b></td>
    </tr>
    <%
    globalWdt = globalWdt + 50


    end if 'export


    '******************************************************************************
    '***** MAIN LOOP DATO INTERVAL ****' 
    '******************************************************************************


     

    call smileyAfslutSettings()

       
    
    for wth = 0 to weekDiff_WeekReal_norm

    

    '*********************************************************
    'cnt = wth

    datoB = dateadd("ww", wth, ddato) '*** ændret fra minus
    datoB = day(datoB)&"-"& month(datoB) &"-"& year(datoB)


    weekDayThisB = weekday(datoB, 2)
    select case weekDayThisB
    case 1
    datoB = dateadd("d", 6, datoB)
    case 2
    datoB = dateadd("d", 5, datoB)
    case 3
    datoB = dateadd("d", 4, datoB)
    case 4
    datoB = dateadd("d", 3, datoB)
    case 5
    datoB = dateadd("d", 2, datoB)
    case 6
    datoB = dateadd("d", 1, datoB)
    case 7
    datoB = dateadd("d", 0, datoB)
    end select
    
    datoA = dateadd("d", -(6), datoB)
    
    '*** Korriger ved månedsskift **'
    if (datepart("m", datoA, 2,2) <> lastMth) AND wth > 0 AND mnow > 12 AND cint(useMDorWeek) = 1 then
    
        datoA = "1-"& month(datoA) &"-"& year(datoA)
        weekDayThisA = weekday(datoA, 2)
        select case weekDayThisA
        case 1
        datoB = dateadd("d", 6, datoA)
        case 2
        datoB = dateadd("d", 5, datoA)
        case 3
        datoB = dateadd("d", 4, datoA)
        case 4
        datoB = dateadd("d", 3, datoA)
        case 5
        datoB = dateadd("d", 2, datoA)
        case 6
        datoB = dateadd("d", 1, datoA)
        case 7
        datoB = dateadd("d", 0, datoA)
        end select

        '** Tilføjer en enkstra uge til loopet
        'weekDiff_WeekReal_norm = weekDiff_WeekReal_norm + 1
        if weekDayThisA <> 1 then
        wth = wth - 1
        end if

    end if    
    'end if



    '** Maks frem til DD ***'
    
    if cint(useMDorWeek) = 1 then 'Skær / Opdel på hele måneder

        if cDate(datoA) < cDate(ddato) then 'ændret <>
        datoA = ddato
        end if

        if mnow <= 12 then
            ugpThisMd = ugp 'Sidste dag i MD eller D.D
        else
    
             '** Hvis der er valgt kvartaler eller ÅTD, skal slut grænsre flytte sig samme med månedsskift
             tjkMth = datepart("m", datoA, 2,2)
             call dageimd(tjkMth, request("yuse"))
             ds = mthDays
             ugpThisMd = ds &"-"& month(datoA) &"-"& year(datoA)

        end if

            if cDate(datoB) > cDate(ugpThisMd) then 'ændret <>
            datoB = ugpThisMd
            end if


            if cDate(datoB) > cDate(now) then 'ændret <>
            datoB = now
            end if

    end if
    
    slutDatoLastm_B = datepart("yyyy", datoB) &"/"& datepart("m", datoB) &"/"& datepart("d", datoB)
    slutDatoLastm_A = datepart("yyyy", datoA) &"/"& datepart("m", datoA) &"/"& datepart("d", datoA)

    'Response.Write "wth: "& wth  &" ugpThisMd: "& ugpThisMd &" ugp: "& ugp &" // weekDiff_WeekReal_norm = "& weekDiff_WeekReal_norm &" Dato A :" &  slutDatoLastm_A & " B:" & slutDatoLastm_B & "<br>"
    

    if media <> "export" then

        if cint(SmiWeekOrMonth) = 1 AND mnow > 12 AND (useSogKriAfs = 0 AND useSogKriGk = 0 AND useSogKri = 0) then 'måned, og kun ved fordeling på ÅTD, Kvartaler mm. og vis ikke totaler når der bruges lukkede uger kriterier

	        if (datepart("m", slutDatoLastm_A, 2,2) <> lastMth) AND wth > 0 then
            call totallinje
            end if

        end if

    end if


            if func = "export" then
	            if last14Mid <> intMids(m) then
                   
                    select case right(m, 1)
                    case 0
                    response.write "<br>"
                    end select

                    if m = 0 then
                    response.write meInit & " (id: " & intMids(m) & ") "
                    else
                    response.write ", "& meInit & " (id: " & intMids(m) & ") " 
                    end if          

                    response.flush
                    last14Mid = intMids(m)
                end if
            end if

    
            call medarbafstem(trim(cint(intMids(m))), slutDatoLastm_A, slutDatoLastm_B, 14, akttype_sel, m)


   
            'Response.Write "<br>"
	        if func <> "export" then
	        response.flush
	        end if
	
	        'end if



        lastWeek = datepart("ww", slutDatoLastm_A, 2,2)
        lastMth = datepart("m", slutDatoLastm_A, 2,2)
        lastslutDatoLastm_A = slutDatoLastm_A



	
	next '**  weeks


 



    if media <> "export" then
	'*********************************************************************************
    '************************* Total *********************************************
    '*********************************************************************************
    call totallinje%>


	

        <%	if media <> "print" AND func <> "export" then %>
        <!--<tr><td colspan="20" align="right">Godkend uger: <input type="checkbox" value="1" id="" /></td></tr>-->
        <%end if %>

    </table>
	
	<%
	end if


    korrektionRealTot = 0
	altimerKorFradTot = 0
	afradragTimerTot = 0
	arealfTimerTot = 0
	anormTimerTot = 0
	atotalTimerPer100 = 0
	anormtime_lontimeTot = 0
	altimerKorFradTot = 0
	arealTimerTot = 0
	aafspTimerTot = 0
	aafspTimerBrTot = 0
	aafspTimerUdbTot = 0
	aafspadUdbBalTot = 0
	aAfspadBalTot = 0
	
	balRealLontimer = 0
	balRealNormtimer = 0

    normtime_lontimeAkk = 0
    balRealNormtimerAkk = 0





        
        flexTimer_tot = 0
        sTimer_tot = 0
        adhocTimer_tot = 0
        ferieAfVal_md_tot = 0 
        ferieAfulonVal_md_tot = 0
        ferieFriAfVal_md_tot = 0
        divfritimer_tot = 0
        omsorg_tot = 0
        tjenestefri_tot = 0

        barsel_tot = 0

        barnSyg_tot = 0
        sygeDage_tot = 0
        lageTimer_tot = 0

    omsorg2afh_tot = 0
    sundTimer_tot = 0
    omsorgKAfh_tot = 0

    showAfsugeTxt_tot = ""
	ugegodkendtTxt_tot = ""

    rejsedage_tot = 0

	end if 'm <> 0 

    lastMid = intMids(m)
	next 'medarb loop
    
     if media <> "export" then%>
	
	
	</td></tr></table>

	
	<%
    end if


           if func = "export" then 'SLUT medarb. load div
    %>
        </div>
    <%
    end if



	if media <> "print" AND func <> "export" then

'Response.flush

ptop = -369
pleft = 1240
pwdt = 200

call eksportogprint(ptop,pleft,pwdt)
%>



    

    <form action="#" method="post">
<tr> 

    <td valign=top align=center>
   <input type=image src="../ill/export1.png" onclick="popUp('week_real_norm_2010.asp?func=export&muse=<%=mnow%>&yuse=<%=year(ugp) %>&FM_moreorless=<%=moreorless %>&FM_saldokri=<%=saldokri %>&FM_timekri=<%=timeKri %>&FM_medarb=<%=thisMiduse%>&FM_md_week=<%=useMDorWeek%>', 400, 200, 200, 100)" />
    </td>
    <td class=lille><input id="Submit3" type="button" value="Eksportér til .csv >> " style="font-size:9px; width:130px;" onclick="popUp('week_real_norm_2010.asp?func=export&muse=<%=mnow%>&yuse=<%=year(ugp) %>&FM_moreorless=<%=moreorless %>&FM_saldokri=<%=saldokri %>&FM_timekri=<%=timeKri %>&FM_medarb=<%=thisMiduse%>&FM_md_week=<%=useMDorWeek%>', 400, 200, 200, 100)" /></td>
</tr>
</form>
<!--
  
    <td align=center><a href="" target="_blank"><img src="../ill/export1.png" border="0"></a></td>
    </td><td><a href="week_real_norm_2010.asp?func=export&muse=<%=mnow%>&yuse=<%=year(ugp) %>&FM_useSogKri=<%=useSogKri%>&FM_moreorless=<%=moreorless %>&FM_saldokri=<%=saldokri %>&FM_timekri=<%=timeKri %>&FM_medarb=<%=thisMiduse%>" class=rmenu target="_blank">.csv fil eksport</a></td>
 
    </tr>

    -->


    <form action="week_real_norm_2010.asp?media=print&muse=<%=mnow%>&yuse=<%=year(ugp) %>&FM_moreorless=<%=moreorless %>&FM_saldokri=<%=saldokri %>&FM_timekri=<%=timeKri %>&FM_medarb=<%=thisMiduse%>&usemrn=<%=session("mid") %>&FM_md_week=<%=useMDorWeek%>" method="post" target="_blank">
<tr>

    <td valign=top align=center>
   <input type=image src="../ill/printer3.png"/>
    </td><td class=lille><input id="Submit6" type="submit" value="Print venlig" style="font-size:9px; width:130px;" /></td>
</tr>
</form>
<!--

    <tr>
    <td align=center><a href="week_real_norm_2010.asp?media=print&muse=<%=mnow%>&yuse=<%=year(ugp) %>&FM_useSogKri=<%=useSogKri%>&FM_moreorless=<%=moreorless %>&FM_saldokri=<%=saldokri %>&FM_timekri=<%=timeKri %>&FM_medarb=<%=thisMiduse%>" target="_blank"  class='rmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
</td><td><a href="week_real_norm_2010.asp?media=print&muse=<%=mnow%>&yuse=<%=year(ugp) %>&FM_useSogKri=<%=useSogKri%>&FM_moreorless=<%=moreorless %>&FM_saldokri=<%=saldokri %>&FM_timekri=<%=timeKri %>&FM_medarb=<%=thisMiduse%>" target="_blank"  class='rmenu'>Print version</a></td>
   </tr>
   
   
  
   -->
   
	
   </table>
</div>
<%else%>

<% if media = "print" then
        Response.Write("<script language=""JavaScript"">window.print();</script>")
   end if
%>
<%end if%>
	
	
	
	<%
	'** Timer Realiseret ***'
	
	'for m = 0 TO UBOUND(intMids)
	'call medarbafstem(intMids(m), startdato, slutdato, 1, akttype_sel)
	'next
	
	%>
	</table>
    
    <% if media <> "print" then   %>
    <!-- table div --> 
	</div>
	<%end if %>


    <form>     <input type="hidden" id="globalWdt" value="<%=globalWdt %>" /></form>
	
	<br /><br />
	
	
	<%if func = "export" then 

    call TimeOutVersion()
   
	ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)
	
	'Response.Write "<br><br>ekspTxt:" & ekspTxt 
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\week_real_norm_2010.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbweekexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbweekexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\medarbweekexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\medarbweekexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "medarbweekexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				
				'Response.end
				
				
				
				'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				'objF.WriteLine(strOskrifter & chr(013))
				objF.WriteLine(ekspTxt)
				
				
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

end if %>

      
      
     
	
	 
	    
	
	<br /><br /><br /><br /><br /><br /><br /><br />
        &nbsp;
			
			</div>
	
	<% 
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
