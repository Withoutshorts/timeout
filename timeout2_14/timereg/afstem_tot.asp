<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->


<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%



 '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "xgrandtotal"

        usemrn = request("jqusemrn")
        startDato = request("jqstartdato")
        slutDato_GT = request("jqslutdato")
        lic_mansat_dt = request("jqlic_mansat_dt")
        levelthis = request("jqlevel")
        

        	call akttyper2009(6)
	        akttype_sel = "#-99#, " & aktiveTyper
	        akttype_sel_len = len(akttype_sel)
	        left_akttype_sel = left(akttype_sel, akttype_sel_len - 2)
	        akttype_sel = left_akttype_sel
        
        'akttype_sel = "#99#"

        'Response.write "her:" & akttype_sel
        'Response.end
        'if session("mid") = 1 then
        'Response.write "startDato: " & startDato & " slutDato_GT: "& slutDato_GT & "<br>"
        'end if

        call medarbafstem(usemrn, startDato, slutDato_GT, 5, akttype_sel, 0)

        lonPerafsl = dateAdd("d", -1, startDato)
        periodeTxt = "<span style=""color:darkred; font-size:11px;"">"& formatdatetime(lic_mansat_dt,1) & " - " & formatdatetime(slutDato_GT, 1) & "<span> "
        periodeTxt = periodeTxt & "<br><span style=""font-size:9px; color:#000000;"">Senest afstemt (l&oslash;n-periode afsluttet): " & formatdatetime(lonPerafsl, 1) & "<span>"
        'if levelthis = 1 then
        'periodeTxt = periodeTxt & " <span style=""color:#999999; font-size:9px; font-weight:lighter;"">(TSA --> Statistik --> Medarb. afstemning & l&oslash;n)</span>"
        'end if

        if len(trim(request("media"))) <> 0 then
        media = request("media")
        else
        media = ""
        end if

        if media = "print" then
        gtDvTop = -220
        else
        gtDvTop = 20
        end if 
        %>
         <div id="gtdv" style="position:relative; top:<%=gtDvTop%>px; left:0px; visibility:visible; display:;">
        <table><tr><td>
        <b>Grandtotal <%=periodeTxt %></b>
          </td></tr></table>

        <table cellspacing=5 cellpadding=2 border=0><tr> 
    <%

    ltf = 0 
    if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then
    %>
    <!-- saldobalancer -->
    <td align=right class=lille style="background-color:pink; border-bottom:2px silver solid; width:80px; white-space:nowrap;"><b>Flekssaldo +/-</b><br /> 
    Realiseret<br /> / Normtid<br />
    <%=realNormBal %>
    </td>
    <%
    ltf = ltf + 160
    end if
    
    
     if session("stempelur") <> 0 then %>
    
     <td align=right class=lille style="background-color:#DCF5BD; border-bottom:2px silver solid; width:80px; white-space:nowrap;"><b>Flekssaldo +/-</b><br /> 
    Komme/g&aring; tid<br /> / Normtid<br />
     <%=normLontBal %>
    </td>

    <%
    
    ltf = ltf + 160
    if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn"  then %>
     <td align=right class=lille style="border-bottom:2px silver solid; width:80px; white-space:nowrap;"><b>Balance +/-</b><br />
    Realiseret<br />/ Komme/g&aring; tid<br />
    <%=realLontBal %>
    </td>
    <%
    ltf = ltf + 160
    end if 
    
    end if
    
    
    
    %>


      <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
           
               
            <%select case lto
               case "tec", "esn"
                overarbTxt = "Afspad."
               case else 
                overarbTxt = "Overarb."
                end select %>


	         <%if lto <> "kejd_pb" AND lto <> "fk" then %>
	         <td align=right style="border-bottom:2px silver solid; width:80px; white-space:nowrap;" valign=bottom class=lille><%=overarbTxt &" optjent:" %><br /> <%=formatnumber(afspTimer(x), 2)%></td>
            <%end if %>
             
            <td align=right class=lille style="border-bottom:2px silver solid; width:80px; white-space:nowrap;;" valign=bottom>Afspaderet:<br /> <%=formatnumber(afspTimerBr(x), 2)%></td>

          
	             <%
                 
                  aafspTimerTot = aafspTimerTot + afspTimer(x) 
	              aafspTimerBrTot = aafspTimerBrTot + afspTimerBr(x)
	              aafspTimerUdbTot = aafspTimerUdbTot + afspTimerUdb(x)
	          
	             afspadUdbBal = 0
	             afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         
	             aafspadUdbBalTot = aafspadUdbBalTot + (afspadUdbBal)

                 AfspadBal = 0 
	             AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	             aAfspadBalTot = aAfspadBalTot + (AfspadBal)
                 
                 if lto <> "lw" AND lto <> "kejd_pb" AND lto <> "fk" then %>
	             <td align=right class=lille style="border-bottom:2px silver solid; width:80px; white-space:nowrap;" valign=bottom>Udbetalt:<br /><%=formatnumber(afspTimerUdb(x), 2)%></td>
	             <td align=right class=lille style="border-bottom:2px silver solid; width:80px; white-space:nowrap;" valign=bottom>&Oslash;nsk. udb.:<br /><%=formatnumber(afspadUdbBal, 2)%></td>
	             <td align=right class=lille style="border-bottom:2px silver solid; width:80px; white-space:nowrap;" valign=bottom>Overarb. tilgode saldo:<br /><%=formatnumber(AfspadBal, 2)%></td>
        	 
                    <%end if %>
	         
            
            

	       <%end if %>

           
            </tr>
             </table>
             </div>

           

        <%

     
           
        end select
        Response.end


        end if  


if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	
	thisfile = "afstem_tot.asp"
	if len(trim(request("usemrn"))) <> 0 then
	usemrn = request("usemrn")
	else
	usemrn = session("mid")
	end if

    level = session("rettigheder")
    
    if len(trim(request("media"))) <> 0 then
    media = request("media")
    else
    media = ""
    end if


      if len(trim(request("medarbsel_form"))) <> 0 then

    if len(trim(request("FM_visAlleMedarb"))) <> 0 then
    visAlleMedarbCHK = "CHECKED"
    visAlleMedarb = 1
    else
    visAlleMedarbCHK = ""
    visAlleMedarb = 0
    end if

    else

        if request.cookies("tsa")("visAlleMedarb") = "1" then
        visAlleMedarbCHK = "CHECKED"
        visAlleMedarb = 1
        else
        visAlleMedarbCHK = ""
        visAlleMedarb = 0
        usemrn = session("mid")
        end if

    end if

    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb


    call medarb_teamlederfor

	

    'if len(trim(request("vis_gt_print"))) <> 0 AND media = "print" then
    'vis_gt_print = 1
    'else
    'vis_gt_print = 0
    'end if

    'if cint(vis_gt_print) = 1 then
    'vis_gt_print_CHK = "CHECKED"
    'else
    'vis_gt_print_CHK = ""
    'end if


	varTjDatoUS_man = request("varTjDatoUS_man")

    if len(trim(request("varTjDatoUS_son"))) <> 0 then
    varTjDatoUS_son = request("varTjDatoUS_son")
    else
    varTjDatoUS_son = dateadd("d", 6, varTjDatoUS_man)
    varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)
	end if

	nextweek = dateadd("d", 7, varTjDatoUS_man)
	prevweek = dateadd("d", -7, varTjDatoUS_man)
	
	'Response.Write "varTjDatoUS_man: " & varTjDatoUS_man
	
	intMid = usemrn

   

    'if len(trim(request("sort"))) <> 0 then
    'sort = request("sort")
    'else
    '       if request.cookies("tsa")("afs_sorter") <> "" then
    '       sort = request.cookies("tsa")("afs_sorter") 
    '       else
    '       sort = 0
    '       end if
    
    'end if

    'response.cookies("tsa")("afs_sorter") = sort


     sort = 1 


    if cint(sort) = 0 then
    fortegn = -1
    else
    fortegn = 1
    end if
                   

 
	

    if media <> "print" AND media <> "export" then


	            %>
	            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

                <%call browsertype()
    
                  

	                call menu_2014() %>
	              <%
    
                 'call treg_3menu(thisfile)

                 toppx = 102
                 bdpx = 1
   
    else
    %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
    <%
    
    bdpx = 0
    toppx = -20
    end if
    %>
    
    
   
   <script src="inc/afstem_jav.js" type="text/javascript"></script>

   
   


      
  
   
   
  
   
   
   <%
   
   if len(trim(request("show"))) <> 0 then
   show = request("show")
    
    if show = 5 then
        'if lto = "cst" OR lto = "xintranet - local" then
        'show = 51
        'else
        show = 12
        'end if
    end if

   else
   show = 12
   end if
   
   'if show = 12 AND instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then 
   'dwdt = 800
   'else

   select case lto
   case "kejd_pb", "cst", "tec", "esn"
   dwdt = 1000
   case else
   dwdt = 1300
   end select
   'end if
   
  
    '**** Indeværende år
    if len(trim(request("yuse"))) = 0 OR cint(request("yuse")) = cint(year(now)) then
    ugp = now
    ddato = datepart("d", ugp) &"/"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
    ddato = dateadd("d",-1, ddato)
    showigartxt = "("&tsa_txt_319&")"

       
        if len(trim(varTjDatoUS_man)) <> 0 then
        yuse = year(varTjDatoUS_man) 'now
        else
        yuse = year(now)
        end if
      
       showfremtidTxt = 0 

    else

       '*** Fremtid
       if cint(request("yuse")) > cint(year(now)) then

        showigartxt = ""
        ugp = "31-01-"&request("yuse")
        ddato = datepart("d", ugp) &"/"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
        yuse = request("yuse")

       showfremtidTxt = 1

       else

        
       '*** Fortid
        showigartxt = ""
        ugp = "31-12-"&request("yuse")
        ddato = datepart("d", ugp) &"/"& datepart("m", ugp) &"/"& datepart("yyyy", ugp)
        yuse = request("yuse")

       showfremtidTxt = 0

       end if

    end if
    
   
    
    
    slutDato = datepart("yyyy", ddato) &"/"& datepart("m", ddato) &"/"& datepart("d", ddato)
    slutDato_GT = dateAdd("d", -1, now)'slutDato
    slutDato_GT = year(slutDato_GT) &"/"& month(slutDato_GT) &"/"& day(slutDato_GT)


       select case lto
       case "kejd_pb", "fk"
       globalWdt = 600
       case "cst"
       globalWdt = 800
       case "tec"

       if show = 77 then
       globalWdt = 1200
       else
       globalWdt = 800
       end if

       case "esn"
       
       if show = 77 then
       globalWdt = 1400
       else
       globalWdt = 900
       end if

       case else 
       
       if show = 77 then
       globalWdt = 1100
       else
       globalWdt = 950
       end if
       
       end select


       if media = "print" then
       sdiv_pd = 0
       else
       sdiv_pd = 20
       end if
  
    %>
   
 
    
    
    <div name="medarbafstem_aar" id="medarbafstem_aar" style="position:relative; left:20px; top:<%=toppx%>px; display:; border:0px #999999 solid; height:auto; visibility:visible; width:<%=globalWdt%>px; background-color:#FFFFFF; padding:<%=sdiv_pd%>px;">
   
        
        <% if media <> "print" AND media <> "export" then %>
	<form id="filterkri" method="post" action="afstem_tot.asp">
        <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
        <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
        <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
        <table>
       <tr bgcolor="#ffffff">
	<td valign=top> <b><%=tsa_txt_077 %>:</b> <br />
    <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)
   
	<br />
				<% 
				call medarb_vaelgandre
                %>
        </td>
           </tr></table>
        </form>
        <br /><br />
	
	<%end if 'media


    '*** Tjekker medarbtpyen **'
    call mtyp_fn(usemrn)
    
  
    if media <> "export" then
    %>
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
	<tr>
	
	<td valign=top><h4>Afstemning (Saldo)<br />
    <span style="font-size:11px; font-weight:lighter;">
     <%call meStamdata(usemrn) %>
    <%=meTxt %>, ansat: <%=formatdatetime(meAnsatDato, 1) %>
    </span>
    </h4>
      

    </td>

	</tr>
 

	</table>
    <%end if %>
   
    <%if media <> "print" AND media <> "export" AND show = 12 then
     %>   
        <div id="gt_msg" style="width:200px; padding:20px; position:absolute; left:120px; background-color:snow; z-index:999999999; border:2px #6CAE1C dashed; display:none; visibility:hidden;">Grandtotal og flexsaldo er flyttet til bunden af deres respektive kolonner på måned/måned oversigten.</div>

        
     <%end if %> 
    
    
    <!--
    Saldo Uge / md - dato:
	 <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=5&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=vmenu><%=global_txt_163 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
     -->
     <%if media <> "print" AND media <> "export" then 
         
         if cint(request("yuse")) = cint(year(now)+1) then
         cls = ""
         else
         cls = "vmenu"
         end if
         %>
  <!--
    <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=12&yuse=<%=year(now)+1%>&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=<%=cls%>> <%=year(now)+1 %></a>&nbsp; 
        -->

    
	<!--<a href="afstem_tot.asp?usemrn=<%=intMid %>&show=12&yuse=<%=year(now)%>&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=vmenu> <%=year(now) %></a>&nbsp; -->
    <%
    

    if month(now) >= 10 then 'Viser næste år
        aarStKri = -1
    else
        aarStKri = 0
    end if
    
    for x = aarStKri to 4
	
	prevyears = dateadd("yyyy", -x, now)
	
         if cint(request("yuse")) = year(prevyears) OR (len(trim(request("yuse"))) = 0 AND x = 0) then
         cls = ""
         YtXT = "<u>" & year(prevyears) & "</u>" 
         else
         cls = "vmenu"
         YtXT = year(prevyears) 
         end if

	%>
	 <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=12&yuse=<%=year(prevyears) %>&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=<%=cls%>><%=YtXT%></a>&nbsp;
	<%
	next %>


    <%
        if datepart("y", now, 2,2) <= 120 then '1 maj
        prevyearsFe = dateadd("yyyy", -2, "1/1/"&year(now)) 
        else
        prevyearsFe = dateadd("yyyy", -1, "1/1/"&year(now)) 
        end if
    
    %>

	|&nbsp;&nbsp;Ferie <%=global_txt_163 %>:
	<a href="afstem_tot.asp?usemrn=<%=intMid %>&show=4&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=vmenu>nuværende</a> 

     /
    <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=4&yuse=<%=year(prevyearsFe) %>&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=gray>forrige ferieår (<%=year(prevyearsFe)%>)</a>&nbsp;
	<br />
	<br />
                <%
    if media <> "print" then
    %>
        <!--
        <table>
        <tr><td style="color:#999999;">  <form>
        <input type="checkbox" name="vis_gt_print" id="vis_gt_print" value="1" <%=vis_gt_print_CHK %> /> Skjul grandtotal
            </form></td></tr>
            </table>
        -->
        <%end if %>
    
    
	
	<div id="loadbar" style="position:absolute; display:; visibility:visible; top:120px; left:250px; width:300px; background-color:#ffffff; border:10px #6CAE1C solid; padding:5px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
       
	</td></tr>
        <tr><td colspan="2" class="lille">Forventet laodtid: 5-20 sek.</td></tr>
	</table>

	</div>

    <%end if%>

        <%if media <> "print" then
            toppx_dvu = 0
          else
            toppx_dvu = -20
           end if  %>


    <div id="dv_udspec" style="position:relative; top:<%=toppx_dvu%>px; left:0px; border:0px #999999 solid;">


	
	
	<%
	
	
	
	'call licensStartDato()
	call meStamdata(usemrn)
	
	ansatdato = meAnsatDato
	mnavn = meTxt 'oRec("mnavn") & " ("& oRec("mnr") &")"
	
	
	call licensStartDato()
	ltoStDato = startDatoDag &"/"& startDatoMd &"/"& startDatoAar
	
	'*** fra ansættelse eller licens startdato ***'
	if cDate(ltoStDato) > cDate(ansatdato) then
	startDato = datepart("yyyy", ltoStDato) &"/"& datepart("m", ltoStDato) &"/"& datepart("d", ltoStDato)
    else
    startDato = datepart("yyyy", ansatdato) &"/"& datepart("m", ansatdato) &"/"& datepart("d", ansatdato)
    end if
    
    
	'dateDiff("ww", slutDato, slutDato, 2, 2) 
	
	'Response.Write startDato  & "ltoStDato:" & ltoStDato & " Weeks:" & dateDiff("ww", startDato, slutDato, 2, 2)  & "<br>"
	'Response.Flush
	
    '** Er egentlig unødvendig at tjekke for da, man allerede er ADMIN eller TEAMleder for at kunne vælge medarbejderen.
	call erTeamLederForvlgtMedarb(level, usemrn, session("mid"))
	
	
	call akttyper2009(6)
	akttype_sel = "#-99#, " & aktiveTyper
	akttype_sel_len = len(akttype_sel)
	left_akttype_sel = left(akttype_sel, akttype_sel_len - 2)
	akttype_sel = left_akttype_sel
	
	select case show 

    
   
    case 12, 77

    showextended = 0
    call stempelur_kolonne(lto, showextended)
    'Response.Write akttype_sel
     



     if show = 12 then
   
    'ddato = "23-4-2013"
    monthThis = month(ddato)
    endfor = monthThis 
    stKri = -1
    endKri = -2
    stfor = stKri 'mth

    visning = 7
    else

   
    monthThis = month(varTjDatoUS_man)
    stKri = 0
    stfor = stKri '0 'datepart("d", ddato, 2,2)
    endKri = -1
    visning = 77
     
    '** Finder altal dage i dag/ dag visning

    

    if (cint(month(now)) = cint(monthThis) AND cint(year(now)) = cint((yuse))) then
        
        endfor = day(ddato)

    else

        select case monthThis
        case 2
        
            select case year(varTjDatoUS_man)
            case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044" 
            endfor = 29
            case else
            endfor = 28
            end select

        case 1,3,5,7,8,10,12
        endfor = 31
        case else
        endfor = 30
        end select


    end if

    end if

    public bgc, lastWeek
    bgc = 0
    lastWeek = 0



   
     
     if media <> "export" then%>
     <b>Udspecificering

      <% 
    if show = 12 then
    Response.Write "År >> Dato</b> &nbsp;"

        
    %>


  
         
         <%if cint(showfremtidTxt) = 1 OR (month(now) = 1 AND day(now) < 10) then %>
         <span style="color:darkred; border:1px red solid; padding:4px;"><b>Bemærk</b> at flekssaldo mv. først stemmer når forrige år er opgjort.</span>
         <%else%>
             
         <span style="color:darkred;">Opgjort pr. <b><%=formatdatetime(ddato, 1) %></b> <%=showigartxt %></span> 
         <%end if %>
         
         <br /><br />&nbsp;
    <%
    
     'if sort <> 0 then
     'sortLnkTxt0 = "<< Dec - Jan"
     'sortLnkTxt1 = "<u>Jan - Dec >></u>"
     'else
     'sortLnkTxt0 = "<u><< Dec - Jan</u>"
     'sortLnkTxt1 = "Jan - Dec >>"
     'end if

    
    else %>
    Måned >> Dato</b>
    <%
    
     'if sort <> 0 then
     'sortLnkTxt0 = "<< 31 - 1"
     'sortLnkTxt1 = "<u>1 - 31 >></u>"
     'else
     'sortLnkTxt0 = "<u><< 31 - 1</u>"
     'sortLnkTxt1 = "1 - 31 >>"
     'end if
    
    end if %>

         <br />

    <%if media <> "print" then %>
   <!-- <br /><a href="Javascript:history.back()" class="vmenu"><< Tilbage</a>-->
        <%end if %>
    

     <table cellpadding=0 cellspacing=0 border=0 width="100%">
     <tr><td valign=top style="padding:0px; border:0px #cccccc solid;">
   

         <%if media <> "print" then 
             pdt = 2
         else
                 
             pdt = 0
         
         end if %>

    <table cellspacing=0 cellpadding=<%=pdt%> border=0 width="100%"><tr>

    <%if show = 77 then %>
         <td style="border-bottom:1px silver solid;">&nbsp;</td>
    <%end if %>
    <td style="border-bottom:1px silver solid;">&nbsp;</td>

	 
     <!-- Norm -->
	
	 <td valign=bottom class=lille style="border-bottom:1px silver solid; width:50px;"><b><%=tsa_txt_173%></b></td>
	 
     
     <!-- Komme / Gå -->
      
	  <%if session("stempelur") <> 0 then 
          %> 
         <%if show = 77 then %>
         <td valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b>Komme/Gå tid</b></td>

                     <%select case lto
                                       case "kejd_pb", "xintranet - local"
                                        case else %>
                    <td valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b>Pauser</b></td>
                    <%end select %>
        

        <%end if %>

      

	         <td valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b>Komme/Gå<br /> optjent</b>

             <%if lto = "dencker" then %>
             <br />(løntimer)
             <%end if %>
     
             </td>


          <%if cint(mtypNoflex) <> 1 then 'noflex %>

                <%if cint(showkgtil) = 1 then
                    globalWdt = globalWdt + 50 %>
	            <td valign=bottom class=lille style="border-bottom:1px silver solid; width:50px;"><b>Tillæg +/-</b><br />Ferie, sygdom, korrektion etc. indberettet</td>
	            <td valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b>= Sum</b><br />Komme/gå tid<br /> + tillæg</td>
                <%end if 
            

       
            
              select case lto 
              case "kejd_pb"
            
               case else%>
	 
	         <td valign=bottom class=lille style="border:1px #DCF5BD solid; border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b><%=tsa_txt_284%> +/-</b>
       
                  <%
                  '*** denne kan slettes for Kejd_pb men vi afventer lige godkendelse af ny opsætningen    
                  select case lto 
                 case "kejd_pb"
       
                 case else
                  %>
                 <br />Komme/gå tid<br /> / Normtid
                 <%end select %>
	         </td>
	         <td bgcolor="#DCF5BD" valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b><%=tsa_txt_284%> +/-<br / >Akkumuleret</b>
         
                 <%select case lto 
                 case "kejd_pb"
       
                 case else
                  %>
                 <br />Komme/gå tid<br /> / Normtid
                 <%end select %>
         
         
                </td>

             <%
             end select    
         
             end if 
         
      end if 'noflex%>


     <!-- Realiseret --->

       <td valign=bottom class=lille style="border-bottom:1px silver solid; width:50px;">
       <%select case lto 
       case "kejd_pb"
       %>
       <b>Timer</b><br /> indberett.<br /> på aktivitet.
       <%
       case else
       %>
       <b>Timer<br /> realiseret</b>
      <br />(indberettet på aktivit.)
       <%
       end select %>
       </td>

        <%if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then %>
	   <td class=lille valign=bottom style="border-bottom:1px silver solid; width:50px;">(heraf<br />fak.bare)</td>
        <%end if %>


         <%if cint(mtypNoflex) <> 1 then 'noflex %>

        <%if cint(showkgtil) = 1 AND (lto <> "cst" AND lto <> "tec" AND lto <> "esn") then %>
	    <td class=lille valign=bottom style="border-bottom:1px silver solid; width:50px;"><b>Korrektion</b><br />(overført saldo fra senest afsluttet løn-periode)</td>
        <%end if %>
    
        

               <%if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then %>
	      <td valign=bottom class=lille style="border:1px pink solid; border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b><%=tsa_txt_284%> +/-</b>
           
              <%select case lto
               case "kejd_pb"%>
              <br /> Indberettet <br />/ Normeret
                <%case else %>  
              <br />Realiseret<br /> / Normeret
	            <%end select %>

              </td>

            <td bgcolor="pink" valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b><%=tsa_txt_284%> +/- <br / >Akkumuleret</b>


                  <%select case lto
               case "kejd_pb"%>
              <br /> Indberettet <br />/ Normeret
                <%case else %>  
              <br />Realiseret<br /> / Normeret</td>
	            <%end select %>

            </td>
          <%end if %>

    


     <%if session("stempelur") <> 0 then %> 
	  <%if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then 
          globalWdt = globalWdt + 50%>
	  <td valign=bottom class=lille style="border-bottom:1px silver solid; white-space:nowrap; width:50px;"><b>Balance +/-</b><br />Realiseret<br />/ Komme/gå tid</td>
	  <%end if %>  
	    
	 <%end if %>
	
      <%end if %>
         
          
	 <!-- Afspad / Overarb --->
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	          <%
                  
               if lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then 
               globalWdt = globalWdt + 100
                  

                   select case lto
                   case "tec", "esn"
                    overarbTxt = "Afspad."
                   case else 
                    overarbTxt = tsa_txt_283
                    end select 
                  
                  
               %>
              <td align=right valign=bottom class=lille style="border-bottom:1px silver solid; width:50px;"><b><%=overarbTxt%><br /> <%=tsa_txt_164%></b><br />(enh.)</td>
	          <%end if %>

              <td valign=bottom style="border-bottom:1px silver solid; white-space:nowrap; width:50px;" class=lille><b>Afspads.</b><br />~ timer</td>
               
               <%if lto <> "lw" AND lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then 
                    globalWdt = globalWdt + 100
                   
                   if lto <> "tec" AND lto <> "esn" then%>
	               <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Udbetalt</b></td>
	               <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Ønsk. Udbe.</b></td>
                    <%end if %>
	          
                
                <td valign=bottom style="border-bottom:1px silver solid;width:50px;" class=lille><b><%=overarbTxt &" "& tsa_txt_280 %></b></td>
               <%end if %>
	       
	 
	 <%end if %>



    <%'TEC / ESN special  

      select case lto
       case "esn", "intranet - local"
        %>
            <!--
           <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_147 %></b> 
         <br /> ~ timer</td>
         -->
           <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_132 %></b> 
         <br /> ~ enh</td>
           <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_157 %></b> 
         <br /> ~ enh</td>

          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_192 %></b> 
         <br /> ~ enh</td>
           <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=left(global_txt_193, 7) & " udbetalt" %></b> 
         <br /> ~ enh</td>

        <%
    end select

     select case lto
       case "esn", "tec"
         
         if visning = 77 then %>
          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_166 %></b> 
         <br /> ~ dage</td>
         <%end if 
          
          %>
         <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_172 %></b> 
         <br /> ~ dage</td>
         
          <%
          if visning = 77 then %>
          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=left(global_txt_166, 26) &" Saldo"%></b> 
          <br /> ~ dage</td>
          <%end if 
             
             
             
          if visning = 77 then  %>
          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_151 %></b> 
         <br /> ~ timer</td>
        <%end if %>

          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_148 %></b> 
         <br /> ~ timer</td>
         
         <%if visning = 77 then  %>
         <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=left(global_txt_151, 28) &" Saldo"%></b> 
         <br /> ~ timer</td>
         <%end if %>
         
         <% if visning = 77 then  %>
         <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_152 %></b> 
         <br /> ~ timer</td>
         <%end if %>

          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_179 %></b> 
         <br /> ~ timer</td>
       
           <% if visning = 77 then  %>
          <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=left(global_txt_152, 28) &" Saldo"%></b> 
         <br /> ~ timer</td>
       <% end if
     end select   
     %>


         <!-- Aldersreduktion --->
	 <%if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then %>
	          <%
                  globalWdt = globalWdt + 200
               %>

               <% if visning = 77 then  %>

                        <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille>
	        <%call akttyper(27, 1) %>
	          <b><%=left(akttypenavn, 10)%>.<br />
              <%=right(akttypenavn, 7)%></b>
	   
	         </td>

              <%end if %>

                <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille>
	     <%call akttyper(28, 1) %>
	     <b><%=left(akttypenavn, 10)%>.<br />
          <%=right(akttypenavn, 5)%></b><br />- timer
	     
	     </td>


         <% if visning = 77 then  %>
       
                 <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille>
	            <%call akttyper(29, 1) %>
	             <b><%=left(akttypenavn, 10)%>.<br />
                  <%=right(akttypenavn, 8)%></b>
	    
	             </td>

          

                 <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille>
	             <%call akttyper(29, 1) %>
	             <b><%=left(akttypenavn, 10)%>.<br />
                  Saldo</b>
	     
	             </td>
         <%end if %>


         <%end if '27/28/29 %>



        
	 
	 
	
	 <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Ferie afholdt</b> 
         <br />
	 ~ dage 
       
      </td>

        <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Ferie afholdt u. løn</b> 
         <br />
	 ~ dage 
       
      </td>

         <%select case lto
             case "tec", "esn"
             ferieFriTxt = "Særligferie"
             case else
             ferieFriTxt = "Feriefridage afholdt"
             end select %>
    
      <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=ferieFriTxt %></b><br />
           ~ dage<br />
           <%select case lto 
             case "kejd_pb", "tec", "esn"
             case else
             %>
            (Incl. udbetalt)
            <%
             end select %>
	 </td>

        <%select case lto
            case "xxintranet - local", "fk" %>
         <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>1 Maj timer</b></td>

         <%end select
             
             select case lto
            case "intranet - local", "adra" 
             globalWdt = globalWdt + 50%>
         <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille>
               <%call akttyper(125, 1) %>
	             <b><%=akttypenavn%></b></td>

         <%end select

             
             select case lto
            case "xxintranet - local", "fk", "kejd_pb" 
             globalWdt = globalWdt + 50%>
        <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Omsorgs<br />dage</b><br />
	 ~ dage<br />
     </td>
    <%end select %>


          <%
          select case lto
            case "xxintranet - local", "fk", "kejd_pb" 
             globalWdt = globalWdt + 50%>
       
        <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b><%=global_txt_179 %></b><br />
	 ~ timer<br />
     </td>
    <%end select %>


        <%
          select case lto
            case "xxintranet - local", "fk", "kejd_pb" 
             globalWdt = globalWdt + 50%>
        <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Barsel/Orlov</b><br />
	 ~ dage<br />
     </td>
    
    <%end select %>


           


          <%
          select case lto
            case "xxintranet - local", "fk" 
             globalWdt = globalWdt + 50%>
       
        <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Læge</b><br />
	 ~ timer<br />
     </td>
    <%end select %>



    	 
	  <%
     


      '** ADMIN, DIG SELV, ELLER LEVEL 1 AND TEAMLEDER
      if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then
          globalWdt = globalWdt + 100 %>
	 <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille>
           <b>Syg</b><br />~ dage</td>
	 <%end if %>

         <%if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then%>
	 <td valign=bottom style="border-bottom:1px silver solid; width:50px;" class=lille><b>Barn syg</b><br />~ dage</td>
	 <%end if %>


     <%if visning = 77 then %>
	 <td valign=bottom style="border-bottom:1px silver solid; width:200px;" class=lille><b>Kommentar</b></td>
	 <%
         
         globalWdt = globalWdt + 200
         end if %>
	 
    </tr>
    <%
    
   
        end if 

   





   
    
    'ddato = "21-1-2013"
    'endfor = 1
    'Response.write "SHWO: "& show &" ddato:" & ddato  &" stfor: " & stfor & "stKri: " & stKri & " endfor+(endKri): "& endfor &" "&endKri & "sort:" & sort &"ddato: "& ddato &" varTjDatoUS_man "& varTjDatoUS_man &"<br>"
   
    'Response.write visning

    
    'StDatoKri77sort = ddato
    if sort = 1 then
    ddato = dateadd("m", -(endfor-1), ddato)
    end if



    '*********************************************'
    '********* Henter AKKu saldo inden Periode ***'
    '*********************************************'
    public akuPreNormLontBal, akuPreRealNormBal 

    
    
    
    '*** Finder startDato for afstemning

         call licensStartDato()
         
         call meStamdata(usemrn)

         '**** StartKri for PreSaldo
         if cdate(meAnsatDato) > cdate(licensstdato) then
         listartDato_GT = meAnsatDato
         else
         listartDato_GT = licensstdato
         end if

      

         if visning = 7 then 'md / md
                if sort <> 1 then
                  
                      if year(now) = year(ddato) then
                      PreSaldoEndDt = "1-1-"&year(ddato)
                      PreSaldoEndDt = dateAdd("d", -1, PreSaldoEndDt)
                      else
                      PreSaldoEndDt = dateAdd("yyyy", -1, ddato)
                      PreSaldoEndDt = "31-12-"&year(PreSaldoEndDt)
                      end if    

                      'Response.write "7 = 1: listartDato_GT: "& listartDato_GT & " ddato:" & ddato &" PreSaldoEndDt: "& PreSaldoEndDt & " lonKorsel_lukketPerDt: "& lonKorsel_lukketPerDt &" endfor:" & endfor &" stKri: "& stKri
                else
                  

                      if year(now) = year(ddato) then
                      PreSaldoEndDt = "1-"&month(ddato)&"-"&year(ddato)
                      PreSaldoEndDt = dateAdd("d", -1, PreSaldoEndDt)
                      else
                      PreSaldoEndDt = dateAdd("yyyy", -1, ddato)
                      PreSaldoEndDt = "31-12-"&year(PreSaldoEndDt)
                      end if    
                 
                      
                     
                 
                      'Response.write "7 <> 1: listartDato_GT: "& listartDato_GT & " PreSaldoEndDt: "& PreSaldoEndDt & " lonKorsel_lukketPerDt: "& lonKorsel_lukketPerDt &" endfor:" & endfor &" stKri: "& stKri
                     'Response.end
                end if
         end if
         
         if visning = 77 then 'dag/dag
            
             if sort <> 1 then
                  
                   StDatoKri77sort = dateadd("d", +(endfor-1-stfor), varTjDatoUS_man)
                  PreSaldoEndDt = "1-"& month(StDatoKri77sort) &"-"&year(StDatoKri77sort)
                  PreSaldoEndDt = dateAdd("d", -1, PreSaldoEndDt)

                   'Response.write "77 <> 1: ddato= "& ddato &" StDatoKri77sort  = "& StDatoKri77sort &" listartDato_GT: "& listartDato_GT & " PreSaldoEndDt: "& PreSaldoEndDt & " lonKorsel_lukketPerDt: "& lonKorsel_lukketPerDt &" endfor:" & endfor &" stKri: "& stKri

            else
                  
                  
                  StDatoKri77sort = dateadd("d", +(stfor), varTjDatoUS_man)
                  PreSaldoEndDt = "1-"&month(StDatoKri77sort)&"-"&year(StDatoKri77sort)
                  PreSaldoEndDt = dateAdd("d", -1, StDatoKri77sort)
                  

                  'Response.write "77 = 1: StDatoKri77sort: "& StDatoKri77sort &" listartDato_GT: "& listartDato_GT & " PreSaldoEndDt: "& PreSaldoEndDt & " lonKorsel_lukketPerDt: "& lonKorsel_lukketPerDt &" endfor:" & endfor &" stKri: "& stKri

                      

            end if
         
         end if
     
                      '** Er der en lukket lønperiode der kan bruges som startdat0
                      call lonKorsel_lukketPerPrDato(PreSaldoEndDt)
                       

                       'if session("mid") = 1 then
                       'Response.write "lonKorsel_lukketPerDt: "& cDate(lonKorsel_lukketPerDt) & " PreSaldoEndDt: "& PreSaldoEndDt & "<br>"
                       'end if

                       '** Den afsluttede lønperiode er før den periode der alsuttes nu
                       '** Den afsluttede lønperiode er senere end licens stdato
                       '** Den afsluttede lønperiode findes (<> 01012002)
                      if ((cDate(lonKorsel_lukketPerDt) <= cDate(PreSaldoEndDt)) AND (cDate(lonKorsel_lukketPerDt) > cDate(listartDato_GT) )) AND cDate(lonKorsel_lukketPerDt) <> "1-1-2002" then
                      lonKorsel_lukketPerDt = dateAdd("d", 1, lonKorsel_lukketPerDt)
                      listartDato_GT = year(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& day(lonKorsel_lukketPerDt)
                      else
                      listartDato_GT = listartDato_GT 
                      end if


                    
                    '*** Henter Saldo før periode ***' 
                   ' Hvis hvert års saldo skal nulstilles ved 31.12, afslut da periode fra HR listen
                   call medarbafstem(usemrn, listartDato_GT, PreSaldoEndDt, 5, akttype_sel, 0)
               
    
    
                '** Omregner til 100 tals
                akuPreNormLontBalTemp = split(normLontBal, ":")
		        
                for t = 0 To UBOUND(akuPreNormLontBalTemp)

                    if t = 0 then
                    akuPreNormLontBal = akuPreNormLontBalTemp(t)
                    else
                    'minuttterne
                    'if session("mid") = 1 AND lto = "esn" then
                    'response.write "normLontBal: "& normLontBal &" akuPreNormLontBalTemp(t) "& akuPreNormLontBalTemp(t) & " = "& formatnumber((akuPreNormLontBalTemp(t) * 100) / 60, 0) &"<br>"
                    'end if        
            
                        select case akuPreNormLontBalTemp(t) 
                        case "00", "01", "02", "03", "04", "05", "06", "07", "08", "09"
                
                            akuPreNormLontBalTemp(t) = cint(akuPreNormLontBalTemp(t))
                            akuPreNormLontBalTemp(t) = formatnumber(akuPreNormLontBalTemp(t) * 100/60, 0)

                        akuPreNormLontBal = akuPreNormLontBal &",0"& akuPreNormLontBalTemp(t) 
        
                        case else
                        akuPreNormLontBal = akuPreNormLontBal &","& formatnumber((akuPreNormLontBalTemp(t) * 100) / 60, 0)
                        end select

                    end if

                next
                '*****
            
              
                'call timerogminutberegning(akuPreNormLontBal) 
                akuPreNormLontBal60 = normLontBal  'akuPreNormLontBal 'thoursTot &":"& left(tminTot, 2) 


    akuPreRealNormBal = realNormBal

    'if session("mid") = 1 AND lto = "esn" then
    'Response.write "normLontBal: "& normLontBal 
    'end if
    '*Slut Pre Saldo 

     
    normLontBal = 0 
    realNormBal = 0
   
    public anormTimerTot, arealTimerTot, atotalTimerPer100, aafspadUdbBalTot, aAfspadBalTot, arealfTimerTot
    public afradragTimerTot, altimerKorFradTot, aafspTimerTot, aafspTimerBrTot, aafspTimerUdbTot
    public ferieAfVal_md, sygDage_barnSyg, ferieAfVal_md_tot, ferieAfulonVal_md_tot, sygDage_barnSyg_tot, ferieFriAfVal_md, ferieFriAfVal_md_tot, normtime_lontimeAkk, balRealNormtimerAkk, korrektionRealTot
    public divfritimer_tot, omsorg_tot, sygeDage_tot, barnSyg_tot, barsel_tot, lageTimer_tot, tjenestefri_tot, rejsedage_tot

    public dagTimer_tot, omsorg2Saldo_tot, natTimer_tot, sundTimer_tot, omsorg10Saldo_tot, weekendTimer_tot, OmsorgKSaldo_tot, omsorgKAfh_tot, omsorg2afh_tot
    public aldersreduktionSaldoTot, aldersreduktionUdbTot, aldersreduktionBrTot, aldersreduktionOpjTot

    public flexTimer_tot, sTimer_tot, adhocTimer_tot, ulempe1706udb_tot, ulempeWudb_tot



    for stfor = stKri to endfor+(endKri)
   
    
       
            if visning = 7 then 'md / md

                if stfor <> -1 OR (cint(endfor) = 1 AND cint(sort) = 1)  then

                    

                    if cint(sort) <> 1 then
                    datoB = dateadd("m", fortegn*stfor, ddato)
                    datoB = "1/"& month(datoB) &"/"& year(datoB)
                    datoB = dateadd("d", fortegn*1, datoB)
                    else

                    '** Sidste loop dd. **' 
                    if stfor = (endfor+(endKri)) then 'sidste loop
                    
                    ''Response.write "sidste loop<br>"
                    datoB = dateadd("m", endfor-1, ddato)
                  
                    else
                    datoB = dateadd("m", fortegn*(stfor+2), ddato)
                    datoB = "1/"& month(datoB) &"/"& year(datoB)
                    datoB = dateadd("d", -1, datoB)
                    end if

                    end if

                    datoA = dateadd("m", fortegn*(stfor+1), ddato)
                
                else 'efter først loop undt sidste loop

                    if cint(sort) <> 1 then
                    datoB = day(ddato)&"/"& month(ddato) &"/"& year(ddato)
                    datoA = ddato
                    else
                    
                  
                    datoB = dateadd("m", fortegn*(stfor+2), ddato)
                    datoB = "1/"& month(datoB) &"/"& year(datoB)
                    datoB = dateadd("d", -1, datoB)

                    datoA = ddato
                    end if
                
                end if

                 

                  slutDatoLastm_B = datepart("yyyy", datoB) &"/"& datepart("m", datoB) &"/"& datepart("d", datoB)
                  slutDatoLastm_A = datepart("yyyy", datoA) &"/"& datepart("m", datoA) &"/1"

                  'Response.write "stfor: "& stfor &" DATOER A: " & slutDatoLastm_A & " og  B: "& slutDatoLastm_B & "<br>"

                  'Response.write "slutDatoLastm_A:" & slutDatoLastm_A & "<br>"
                  'Response.write "slutDatoLastm_B:" & slutDatoLastm_B & "<hr>" 

            else '77 dag / dag

          
            if cint(sort) <> 1 then
            datoB = dateadd("d", +(endfor-1-stfor), varTjDatoUS_man)
            datoA = datoB
            else
            datoB = dateadd("d", +(stfor), varTjDatoUS_man)
            datoA = datoB
            end if
            
                

                  slutDatoLastm_B = datepart("yyyy", datoB) &"/"& datepart("m", datoB) &"/"& datepart("d", datoB)
                  slutDatoLastm_A = slutDatoLastm_B

            end if
        

    
    
  
    
    
    'Response.Write slutDatoLastm_A & " - " & slutDatoLastm_B & "akttype_sel: "& akttype_sel
     
  
    call medarbafstem(usemrn, slutDatoLastm_A, slutDatoLastm_B, visning, akttype_sel, 0)
    'Response.Write "<br>"
	response.flush
	
	next

    
    if media <> "export" then

    '**********************************************
    '************** Total *************************
    '**********************************************
	%>

    <!-- Total -->
	<tr bgcolor="#CCCCCC">

                <%if show = 77 then
                cpt = 2
                else
                cpt = 1
                end if %>    

	<td class=lille style="white-space:nowrap;" colspan="<%=cpt %>"><b>Total i per.</b></td>

        
	 <%if session("stempelur") <> 0 then 
              call timerogminutberegning(anormTimerTot*60)
		                anormTimerTotShow = ""& thoursTot &":"& left(tminTot, 2)
               else
                        anormTimerTotShow = formatnumber(anormTimerTot,2)
         end if%>
	
    <td align=right style="white-space:nowrap;" class=lille><b><%=anormTimerTotShow%></b></td>
	
	<%if session("stempelur") <> 0 then %> 
            
        <%if show = 77 then %>
              <td align=right style="white-space:nowrap;" class=lille>
	        &nbsp;
	        </td>
         <%select case lto
        case "kejd_pb", "xxintranet - local"
        case else %>
         <td align=right style="white-space:nowrap;" class=lille>
	        &nbsp;
	        </td>
        <%end select %>
        <%end if %>

	         <td align=right style="white-space:nowrap;" class=lille>
	         <%call timerogminutberegning(atotalTimerPer100) %>
            <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	        </td>


        <%if cint(mtypNoflex) <> 1 then 'noflex %>

            <%select case lto
             case "kejd_pb", "xxintranet - local"
              case else %>

                 <%if cint(showkgtil) = 1 then %>
	             <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber(afradragTimerTot, 2) %></b></td>
	             <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber(altimerKorFradTot, 2) %></b></td>
	             <%end if
	 
                 %> 
	             <td align=right style="white-space:nowrap;" class=lille>
	             <%

                 anormtime_lontimeTot = -((anormTimerTot - (altimerKorFradTot)) * 60) 
     
                 call timerogminutberegning(anormtime_lontimeTot) %>
		            <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	             </td>

                 <% call timerogminutberegning(normtime_lontimeAkk + (akuPreNormLontBal * 60))  %>
                      <td align=right style="white-space:nowrap;" class=lille>
	
		                <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	                 </td>

            <%end select %>

  
            <%end if %>


        <%end if 'noflex %>

             <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber(arealTimerTot,2)%></b></td>
	         
            <%if lto <> "cst" ANd lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then %>
	        <td align=right style="white-space:nowrap;" class=lille>(<%=formatnumber(arealfTimerTot,2)%>)</td>
            <%end if %>



        <%if cint(mtypNoflex) <> 1 then 'noflex %>

                <%if cint(showkgtil) = 1 AND (lto <> "cst" AND lto <> "tec" AND lto <> "esn") then %>
                <td align=right style="white-space:nowrap;" class=lille><%=formatnumber(korrektionRealTot,2)%></td>
                <%end if %>

	        


	    <%if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then %>
	    <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber((arealTimerTot - anormTimerTot),2)%></b></td>
        <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber(balRealNormtimerAkk+(akuPreRealNormBal),2)%></b></td>
         <%end if %>


           <%if session("stempelur") <> 0 AND (lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn") then %>
        <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber((arealTimerTot - (altimerKorFradTot)),2)%></b></td>
	    <%end if %>

       
        <%end if %>


	 
	    
	    
	    
	      <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	          
             <%if lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then   %> 
	         <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber(aafspTimerTot, 2)%></b></td>
             <%end if %>
             
	         <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aafspTimerBrTot, 2)%></b></td>

	         <%if lto <> "lw" AND lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then %>

                <%if lto <> "tec" AND lto <> "esn" then %>
                <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aafspTimerUdbTot, 2)%></b></td>
	            <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aafspadUdbBalTot, 2)%></b></td>
                <%end if %>
                <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aAfspadBalTot, 2)%></b></td>
                <%end if %>
	         
	      <%end if %>



       <%'TEC/ESN special  

       select case lto
       case "esn", "intranet - local"
        %>

            <!--
           <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(flexTimer_tot, 2) %></b></td>
            -->
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(sTimer_tot, 2) %></b></td>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(adhocTimer_tot, 2) %></b></td>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(ulempe1706udb_tot, 2) %></b></td>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(ulempeWudb_tot, 2) %></b></td>

        <%
    end select


     select case lto
       case "esn", "tec"
         if show = 77 then%>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(dagTimer_tot, 2) %></b></td>
        <%end if %>
        
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(omsorg2afh_tot, 2) %></b></td>
        
        <%if show = 77 then %>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(omsorg2Saldo_tot, 2)%></b></td>
        <%end if %>

        <%if show = 77 then %>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(natTimer_tot, 2)%></b></td>
        <%end if %>
        
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(sundTimer_tot, 2) %></b></td>
        
        <%if show = 77 then %>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(omsorg10Saldo_tot, 2)%></b></td>
        <%end if %>

        <%if show = 77 then %>
          <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(weekendTimer_tot, 2) %></b></td>
        <%end if %>

        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(omsorgKAfh_tot, 2) %></b></td>
        
        <%if show = 77 then %>
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(OmsorgKSaldo_tot, 2)%></b></td>
        <%end if %>
        <%
     end select   
     %>

       <%if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then %>

        <%if show = 77 then%>
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aldersreduktionOpjTot, 2)%></b></td>
        <%end if %>
        
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aldersreduktionBrTot, 2)%></b></td>
        
         <%if show = 77 then%>
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aldersreduktionUdbTot, 2)%></b></td>
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(aldersreduktionSaldoTot, 2)%></b></td>
        <%end if %>
    

       <%end if %>
	  

       <td align=right  class=lille>
       <%if media <> "print" AND media <> "export" then %>
       <a href="afstem_tot.asp?usemrn=<%=usemrn %>&show=4&varTjDatoUS_man=<%=varTjDatoUS_man %>" class=rmenu><%=formatnumber(ferieAfVal_md_tot, 2) %></a>
       <%else %>
       <b><%=formatnumber(ferieAfVal_md_tot, 2) %></b>
       <%end if %>
     </td>
	

         <td align=right  class=lille><b><%=formatnumber(ferieAfulonVal_md_tot, 2) %></b></td>

	 <td align=right  class=lille><b><%=formatnumber(ferieFriAfVal_md_tot, 2) %></b></td>


        <%   select case lto
                        case "xintranet - local", "fk"
            %>
        	 <td align=right  class=lille><b><%=formatnumber(divfritimer_tot, 2) %></b></td>

        <%
            end select %>

          <% 'Rejsedage 
            select case lto
            case "intranet - local", "adra"
            %>
        	 <td align=right  class=lille><b><%=formatnumber(rejsedage_tot, 2) %></b></td>

            <%
            end select %>

        
        <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille><b><%=formatnumber(omsorg_tot, 2) %></b></td>

        <%
            end select %>

           <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille><b><%=formatnumber(tjenestefri_tot, 2) %></b></td>
        

        <%
            end select %>


         <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille><b><%=formatnumber(barsel_tot, 2) %></b></td>
        

        <%
            end select %>


              


            <%   select case lto
                        case "xxintranet - local", "fk"
            %>
        
        <td align=right  class=lille><b><%=formatnumber(lageTimer_tot, 2) %></b></td>

        <%
            end select %>



	 
	  <%if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then%>
	 <td align=right  class=lille><b><%=formatnumber(sygeDage_tot, 2) %></b></td>
	 <%end if %>

          <%if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then%>
	 <td align=right  class=lille><b><%=formatnumber(barnSyg_tot, 2) %></b></td>
	 <%end if %>


           <%if visning = 77 then %>
         <td align=right  class=lille>&nbsp;</td>
        <%end if %>
	
	</tr>
    
    <%response.flush 








   '****************************************************
   '********************** Grandtotal ******************
    '****************************************************
   %>


     <form>
        <!--
    <input type="hidden" id="jqstartdatoDK" value="<%=day(datoA)&"/"&month(datoA)&"/"&year(datoA) %>" />
    <input type="hidden" id="jqstartdato" value="<%=listartDato_GT %>" />
    <input type="hidden" id="jqlic_mansat_dt" value="<%=lic_mansat_dt %>" />
    <input type="hidden" id="jqslutdato" value="<%=slutDato_GT %>" />
    <input type="hidden" id="jqusemrn" value="<%=usemrn %>" />
    <input type="hidden" id="jqlevel" value="<%=level %>" />
            -->

    <input type="hidden" id="show" value="<%=show %>" />
    <input type="hidden" id="media" value="<%=media %>" />
    <input type="hidden" id="globalWdt" value="<%=globalWdt %>" />
 
    </form>


    <%


        if show = 12 then
 
            
            
 '*** Finder startDato for afstemning

 call licensStartDato()
 call lonKorsel_lukketPer(now, -2)
 call meStamdata(usemrn)

 if cdate(meAnsatDato) > cdate(licensstdato) then
 listartDato_GT = meAnsatDato
 else
 listartDato_GT = licensstdato
 end if

 lic_mansat_dt = listartDato_GT

 if cdate(lonKorsel_lukketPerDt) > cdate(listartDato_GT) then
 lonKorsel_lukketPerDt = dateAdd("d", 1, lonKorsel_lukketPerDt)
 listartDato_GT = year(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& day(lonKorsel_lukketPerDt)
 else
 listartDato_GT = year(listartDato_GT) &"/"& month(listartDato_GT) &"/"& day(listartDato_GT)
 end if

         'response.write "Yuse" & yuse

       

        'usemrn = request("jqusemrn")
        'startDato = request("jqstartdato")
        'slutDato_GT = request("jqslutdato")
        'lic_mansat_dt = request("jqlic_mansat_dt")
        'levelthis = request("jqlevel")
        

        	call akttyper2009(6)
	        akttype_sel = "#-99#, " & aktiveTyper
	        akttype_sel_len = len(akttype_sel)
	        left_akttype_sel = left(akttype_sel, akttype_sel_len - 2)
	        akttype_sel = left_akttype_sel


            call medarbafstem(usemrn, listartDato_GT, slutDato_GT, 5, akttype_sel, 0)
            


        lonPerafsl = dateAdd("d", -1, listartDato_GT)
        periodeTxt = "<span style=""font-size:9px; color:#000000;"">Grandtotal periode: "& formatdatetime(lic_mansat_dt,1) & " - " & formatdatetime(slutDato_GT, 1) 
        periodeTxt = periodeTxt & "<br>Senest afstemt (l&oslash;n-periode afsluttet): " & formatdatetime(lonPerafsl, 1) & "<span>"
        
         %>
 
	<tr bgcolor="#FFFFFF">
    <td class=lille style="white-space:nowrap;" colspan="<%=cpt %>"><b>Grandtotal</b></td>
	
	
	<!-- Normtid -->
    <td align=right style="white-space:nowrap;" class=lille>&nbsp;</td>
	
	<%if session("stempelur") <> 0 then %> 
            
        <%if show = 77 then %>
              <td align=right style="white-space:nowrap;" class=lille>
	        &nbsp;
	        </td>
         
        <%select case lto
        case "kejd_pb", "xintranet - local"
        case else %>
         <td align=right style="white-space:nowrap;" class=lille>
	        &nbsp;
	        </td>
        <%end select %>
        <%end if %>

	         <td align=right style="white-space:nowrap;" class=lille>
	       &nbsp;
	        </td>


        <%if cint(mtypNoflex) <> 1 then 'noflex %>

             <%if cint(showkgtil) = 1 then %>
	         <td align=right style="white-space:nowrap;" class=lille> &nbsp;</td>
	         <td align=right style="white-space:nowrap;" class=lille> &nbsp;</td>
	         <%end if
	 
             select case lto
             case "kejd_pb"
             case else %>

            
	         <td align=right style="white-space:nowrap;" class=lille>
	         &nbsp;
	         </td>

     
            <td align=right style="white-space:nowrap; background-color:#DCF5BD;" class=lille><b><%=normLontBal %></b></td>

            <%end select %>

  
  <%end if %>
        <%end if 'noflex %>

             <td align=right style="white-space:nowrap;" class=lille>&nbsp;</td>
	         

         

            <%'Saldo Norm/Real
            if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then %>
	        <td align=right style="white-space:nowrap;" class=lille>&nbsp;</td>
            <%end if 


        if cint(mtypNoflex) <> 1 then 'noflex


                if cint(showkgtil) = 1 AND (lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn") then %>
                <td align=right style="white-space:nowrap;" class=lille>&nbsp;</td>
                <%end if 
        
                         if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then %>
	                    <td align=right style="white-space:nowrap;" class=lille>&nbsp;</td>
                        <td align=right style="white-space:nowrap;" class=lille bgcolor="pink"><b><%=formatnumber(realNormBal,2)%></b></td>

                                   <%if session("stempelur") <> 0 AND (lto <> "kejd_pb") then %>
                                <td align=right style="white-space:nowrap;" class=lille><b> <%=realLontBal %></b></td>
	                            <%end if %>

                        <%end if %>

        <%end if 'noflex%>
	 
	    
	    
	    
	      <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	          
             <%if lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then   %> 
	         <td align=right style="white-space:nowrap;" class=lille><b><%=formatnumber(afspTimer(x), 2)%></b></td>
             <%end if %>
             
            <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(afspTimerBr(x), 2)%></b></td>

             <%
                 
                  aafspTimerTot = aafspTimerTot + afspTimer(x) 
	              aafspTimerBrTot = aafspTimerBrTot + afspTimerBr(x)
	              aafspTimerUdbTot = aafspTimerUdbTot + afspTimerUdb(x)
	          
	             afspadUdbBal = 0
	             afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         
	             aafspadUdbBalTot = aafspadUdbBalTot + (afspadUdbBal)

                 AfspadBal = 0 
	             AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	             aAfspadBalTot = aAfspadBalTot + (AfspadBal)
                %>


	         <%if lto <> "lw" AND lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then %>

                <%if lto <> "tec" AND lto <> "esn" then %>
                <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(afspTimerUdb(x), 2)%></b></td>
	            <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(afspadUdbBal, 2)%></b></td>
                <%end if %>
                <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(AfspadBal, 2)%></b></td>
                <%end if %>
	         
	          
        	 
	        <%end if %>




         <%'TEC/ESN special  

             select case lto
               case "esn", "intranet - local"
             %>
            <!-- Ulempe -->
            <td >&nbsp;</td>
            <td >&nbsp;</td>
            <td >&nbsp;</td>
             <td >&nbsp;</td>
            <td >&nbsp;</td>
              <%
             end select

     select case lto
       case "esn", "tec"
        if visning = 77 then%>
          <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <%end if %>
        <td >&nbsp;</td>
         <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(sundtimer(x), 2) %></b></td>
        <td align=right class=lille style="white-space:nowrap;"><b><%=formatnumber(tjenestefri(x), 2) %></b></td>
          
       <%
     end select   



             if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then 

     %> 
        
        <%if show = 77 then%>
        <td >&nbsp;</td>
        <%end if %>
        
        
        <td >&nbsp;</td>

        <%if show = 77 then%>
        <td >&nbsp;</td>
        <td >&nbsp;</td>
        <%end if %>


        <%end if %>

        <!--Ferie dage afholdt -->
       <td align=right  class=lille>&nbsp;</td>
	
         <!--Ferie dage afholdt u løn -->
	 <td align=right  class=lille>&nbsp;</td>

         <!--Ferie fridage afholdt -->

        <%   select case lto
                        case "xintranet - local", "fk"
            %>
        	 <td align=right  class=lille>&nbsp;</td>

        <%
            end select %>

          <!-- 1 maj timer -->
        <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>

        	 <td align=right  class=lille>&nbsp;</td>

        <%
            end select %>


          <% 'Rejsedage 
            select case lto
            case "intranet - local", "adra"
            %>
        	 <td align=right  class=lille>&nbsp;</td>

            <%
            end select %>

         <!-- Omsorgsdage -->
         <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille>&nbsp;</td>
       
        <%
            end select %>


               <!-- Tjenestefri -->
           <%   select case lto
                        case "xintranet - local", "fk"
            %>
        	 <td align=right  class=lille>&nbsp;</td>
       

        <%
            end select %>

             <!-- Barsel / orlov -->
           <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille>&nbsp;</td>
       
        <%
            end select %>


       
     

            <!-- Læge -->
           <%   select case lto
                        case "xxintranet - local", "fk"
            %>
        	 <td align=right  class=lille>&nbsp;</td>
       

        <%
            end select %>

	 
         <!-- Syg -->
	  <%if level = 1 OR (session("mid") = usemrn) then %>
	 <td align=right  class=lille>&nbsp;</td>
	 <%end if %>

         <!-- barn syg -->
          <%if level = 1 OR (session("mid") = usemrn) then %>
	 <td align=right  class=lille>&nbsp;</td>
	 <%end if %>


         <!-- kommentar -->

             <%if visning = 77 then %>
         <td align=right  class=lille>&nbsp;</td>
        <%end if %>
	
	</tr>
    
    <%response.flush %>

    <tr><td colspan="40">
       
        <%=periodeTxt %>
        </td>




    </tr>

    
    


   

            <%if media <> "print" then %>

             
                <tr><td style="color:#000000; font-size:9px;" colspan="40">>> Saldo videreført fra tidligere periode<br />
                    <!--<img src="ill/blank.gif" width="1140" height="2" border="1" />-->

                     <br />
                Ferie afholdt er incl. ferie udbetalt.
                    </td></tr>
      
                <!--Hvis overført saldo = 0, er gl. saldo evt. overført som korrektion i den nye periode.--> </td></tr>
                <tr><td colspan=40 style="color:#999999; font-size:9px;"><br /><br />
             <b>Regler for startdato for periode til afstemning af grandtotal</b><br /><br />
                         1) <b>Licensstartdato</b><br />
                         2) <b>Medarbejder ans&aelig;ttelses-dato</b>, hvis den er nyere end Licensstartdato.<br />
                         3) <b>Lønperiode afsluttetdato</b>, hvis den er nyere end medarbejder ans&aelig;ttelses-dato. 
            </td></tr>
        

                
         <%end if %>

     <%end if 'vis = 12%>

    </table>

	
	</td></tr></table>



       

   

   <%
   end if


    case 4
    
    '** Ferie / Ferie år + feriefridag **'
    %>
     <table cellpadding=0 cellspacing=0 border=0 width=80%>
     <tr><td valign=top style="padding:20px; border:0px #cccccc solid;">
    <%
    strAar = year(ugp) 'strAar 'year(now)
    
    '*** Ferie år ***

     select case lto

     case "akelius", "intranet - local" 

        ferieaarST = strAar &"/1/1"
	    ferieaarSL = strAar &"/12/31"

    case else

	    if cdate(ugp) >= cdate("1-5-"& strAar) AND cdate(ugp) <= cdate("31-12-"& strAar) then
	    'Response.Write "OK her"
	    ferieaarST = strAar &"/5/1"
	    ferieaarSL = strAar+1 &"/4/30"
	    else
	    ferieaarST = strAar-1 &"/5/1"
	    ferieaarSL = strAar &"/4/30"
	    end if
	

    end select

    'Response.Write "Periode: "& formatdatetime(ferieaarST, 1) &" - "& formatdatetime(ferieaarSL, 1)
    
    '** er ferie slået til ***
    aty_on = 0
    call akttyper2009Prop(14)
    if aty_on = 1 then
        akttype_sel = "#-99#, #11#, #14#, #15#, #16#, #19#, "
        
        aty_on = 0
        call akttyper2009Prop(111)
        if aty_on = 1 then
        akttype_sel = akttype_sel & "#111#, "
        end if

        aty_on = 0
        call akttyper2009Prop(112)
        if aty_on = 1 then
        akttype_sel = akttype_sel & "#112#, "
        end if
        

    
    else
    akttype_sel = "#-99#, "
    end if

    '** er feriefridage slået til ***
    aty_on = 0
    call akttyper2009Prop(13)
    if aty_on = 1 then
    akttype_sel = akttype_sel & "#12#, #13#, #17#, #18#, "
    else
    akttype_sel = akttype_sel
    end if

    
    call medarbafstem(usemrn, ferieaarST, ferieaarSL, 4, akttype_sel, 0)


    
    %>
    </td></tr>
    <tr><td>

        <br /><br />
          <table cellpadding="20" cellspacing="1" border="0"><tr><td valign="top">

            <table cellpadding="0" cellspacing="0" border="0" width="250">
            <tr><td colspan="2"><b>Udspecificering ferie:</b></td></tr>
        <%
            '**Udspecificering
            strSQLfe = "SELECT tdato, timer FROM timer WHERE tmnr = "& usemrn &" AND tdato BETWEEN '"& ferieaarST &"' AND '"& ferieaarSL &"' AND tfaktim = 14" 
            
            'Response.write strSQLfe
            oRec.open strSQLfe, oConn, 3
            while not oRec.EOF 

            %>
            <tr><td align="right" style="border-bottom:1px #cccccc solid;"><%=formatdatetime(oRec("tdato"), 1) %></td><td align="right" style="border-bottom:1px #cccccc solid;"><%=oRec("timer") &" t." %></td></tr>

            <%
            oRec.movenext
            wend
            oRec.close

            %>

            </table>
       

         <% if aty_on = 1 then %>
         </td><td valign="top">

        

           
        <table cellpadding="0" cellspacing="0" border="0" width="250">
            <tr><td colspan="2"><b>Udspecificering feriefridage:</b></td></tr>
        <%
            '**Udspecificering
            strSQLfe = "SELECT tdato, timer FROM timer WHERE tmnr = "& usemrn &" AND tdato BETWEEN '"& ferieaarST &"' AND '"& ferieaarSL &"' AND tfaktim = 13" 
            
            'Response.write strSQLfe
            oRec.open strSQLfe, oConn, 3
            while not oRec.EOF 

            %>
            <tr><td align="right" style="border-bottom:1px #cccccc solid;"><%=formatdatetime(oRec("tdato"), 1) %></td><td align="right" style="border-bottom:1px #cccccc solid;"><%=oRec("timer") &" t." %></td></tr>

            <%
            oRec.movenext
            wend
            oRec.close

            %>

            </table>
       


         <%end if %>
     </td></tr>
     </table>

             </td></tr></table>

   

    <%
    end select%>


    </div> <!-- 'dv_udspec -->
        
         </div> <!-- side div -->


<%if media <> "print" AND media <> "export" then %>
<br /><br /><br /><br /><br /><br /><br /><br />&nbsp
<%end if %>








      <%if media = "export" then 


                if show = 12 OR show = 77 then
               
                strEkspHeader = "Medarbejder;Medarb. nr;Initialer;Dato;"

                strEkspHeader = strEkspHeader & tsa_txt_173 & ";" 'Norm
    
               
                if session("stempelur") <> 0 then
                        if show = 77 then
                        strEkspHeader = strEkspHeader & "Komme/gå tid;"

                          select case lto
                           case "kejd_pb", "xintranet - local"
                            case else 
                        strEkspHeader = strEkspHeader & "Pauser;"
                        end select
                        end if
                
                strEkspHeader = strEkspHeader & "Komme/gå optjent;"
              

                if cint(mtypNoflex) <> 1 then 'noflex 

                    if cint(showkgtil) = 1 then
                    strEkspHeader = strEkspHeader & "Tillæg +/- (ferie, syg mv.);"
                    strEkspHeader = strEkspHeader & "Sum;"
                    end if

                    
                        select case lto
                        case "kejd_pb"
                        case else     
                        strEkspHeader = strEkspHeader & tsa_txt_284 &" +/- (Komme/gå tid / Normtid);" 'Saldo
                        strEkspHeader = strEkspHeader & tsa_txt_284 &" +/- Akkumuleret (Komme/gå tid / Normtid);" 'Saldo akkumuleret
                        end select
                    end if
                
                end if

              
                select case lto
                case "kejd_pb"
                strEkspHeader = strEkspHeader & "Timer indberettet på aktiviteter;"
                case else
                strEkspHeader = strEkspHeader & "Timer realiseret;"
                end select


                if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then
                strEkspHeader = strEkspHeader & "(heraf fakturerbare);"
                  end if  
          

             if cint(mtypNoflex) <> 1 then 'noflex 

                 if cint(showkgtil) = 1 AND lto <> "tec" AND lto <> "esn" then
                   strEkspHeader = strEkspHeader & "Korrektion (overført saldo);"
                 end if

              

                  if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then 

                     select case lto
                    case "kejd_pb"
                    strEkspHeader = strEkspHeader & tsa_txt_284 &" +/- (Indberettet / Normtid);" 'Saldo
                    strEkspHeader = strEkspHeader & tsa_txt_284 &" +/- Akkumuleret (Indberettet / Normtid);" 'Saldo
                    case else   
                    strEkspHeader = strEkspHeader & tsa_txt_284 &" +/- (Realiseret / Normtid);" 'Saldo
                    strEkspHeader = strEkspHeader & tsa_txt_284 &" +/- Akkumuleret (Realiseret / Normtid);" 'Saldo
                    end select

                end if
               



               if session("stempelur") <> 0 then 
                   

                    if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then
                    strEkspHeader = strEkspHeader & tsa_txt_166 &" +/- (Realiseret / Komme/gå tid (løntimer);" 'balance
                    end if

               end if


               end if 'noflex

       
               if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then
               
               if lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" then
               strEkspHeader = strEkspHeader & tsa_txt_283 &" "& tsa_txt_164 &" (enh.);"
               end if

               strEkspHeader = strEkspHeader & "Afspads. ~ timer;"

                if lto <> "lw" AND lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" then
                    if lto <> "tec" AND lto <> "esn" then
                    strEkspHeader = strEkspHeader &"Udbetalt;Ønsket udb.;"
                    end if
                
                
                strEkspHeader = strEkspHeader & tsa_txt_283 &" "& tsa_txt_280 &";"
                                
                end if

               end if


               'TEC / ESN special  
                select case lto
                case "esn", "intranet - local"
                'strEkspHeader = strEkspHeader & global_txt_147&";"& global_txt_132&";"& global_txt_157&";"& global_txt_192&";"& global_txt_193&";"
                strEkspHeader = strEkspHeader & global_txt_132&";"& global_txt_157&";"& global_txt_192&";"& global_txt_193&";"
                end select

                select case lto
                case "esn", "tec"
                if visning = 77 then
                strEkspHeader = strEkspHeader &"Omsorg 2 optj.;Omsorg 2 afh.;Omsorg 2 saldo;Omsorg 10 optj.;Omsorg 10 afh.;Omsorg 10 saldo;Omsorg konverteret optj.;Omsorg konverteret afh.;Omsorg konverteret saldo;"
                else
                strEkspHeader = strEkspHeader &"Omsorg 2 afholdt;Omsorg 10 afholdt;Omsorg konverteret afholdt;"
                end if
                end select   

                
                 if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then
                        
                        if visning = 77 then
                        strEkspHeader = strEkspHeader &"Aldersreduktion Optjent;"
                        end if

                          
                        strEkspHeader = strEkspHeader &"Aldersreduktion Brugt;"
                      
                        if visning = 77 then
                        strEkspHeader = strEkspHeader &"Aldersreduktion Udbetalt;Aldersreduktion Saldo;"
                        end if

                end if

     

                  select case lto
                 case "tec", "esn"
                 ferieFriTxt = "Særligferie"
                 case else
                 ferieFriTxt = "Feriefridage afholdt"
                 end select

               strEkspHeader = strEkspHeader &"Ferie afholdt ~ dage;Ferie afholdt u. løn ~ dage;"& ferieFriTxt &" ~ dage;"

                
                select case lto
                case "xintranet - local", "fk" 
                 strEkspHeader = strEkspHeader &"1 maj timer;"
                end select
                

                select case lto
                case "xintranet - local", "fk", "kejd_pb" 
                 strEkspHeader = strEkspHeader &"Omsorgsdage ~ dage;"
                end select

                 select case lto
                case "xintranet - local", "fk", "kejd_pb" 
                 strEkspHeader = strEkspHeader &"Tjenestefri ~ timer;"
                end select

                select case lto
                case "xintranet - local", "fk", "kejd_pb" 
                 strEkspHeader = strEkspHeader &"Barsel/Orlov ~ dage;"
                end select

            

                select case lto
                case "xxintranet - local", "fk" 
                 strEkspHeader = strEkspHeader &"Læge ~ timer;"
                end select



               if session("rettigheder") = 1 OR (session("mid") = usemrn) then
               strEkspHeader = strEkspHeader & "Syg ~ dage;"
               end if

               if session("rettigheder") = 1 OR (session("mid") = usemrn) then
               strEkspHeader = strEkspHeader & "Barn syg ~ dage;"
               end if
            
               if visning = 77 then 
                strEkspHeader = strEkspHeader & "kommentar;"
               end if

                   
               strEkspHeader = strEkspHeader & "xx99123sy#z"
                
               else '*** Feriefri og Ferie
                   
                   if (lto <> "cst" AND lto <> "tec" AND lto <> "esn") AND instr(akttype_sel, "#13#") <> 0 then 
                   strEkspHeaderA = "Feriefridage ("&ferieFriaarStart&")" & "xx99123sy#z"
                   strEkspHeaderA = strEkspHeaderA & meNavn & ";" & meNr & ";"& meInit & ";" & "xx99123sy#z"

                  
                   strEkspHeaderA = strEkspHeaderA & tsa_txt_174 &" "& tsa_txt_164 &" ~ dage;Planlagt >> dd.~ dage;"&tsa_txt_165 &"~ dage;;Udbetalt ~ dage;"&tsa_txt_282 &" "& tsa_txt_280 &" ~ dage;"

                   strEkspHeaderA = strEkspHeaderA & "xx99123sy#z"
                   end if


                   if instr(akttype_sel, "#14#") <> 0 then 
                   
                       select case lto
                       case "akelius", "intranet - local"
                       strEkspHeaderB = tsa_txt_281 &" (1.1."&ferieaarStart&" - 31.12."&ferieaarSlut&")" & "xx99123sy#z"
                       case else
                       strEkspHeaderB = tsa_txt_281 &" (1.5."&ferieaarStart&" - 30.4."&ferieaarSlut&")" & "xx99123sy#z"
                       end select
                   
                   strEkspHeaderB = strEkspHeaderB & meNavn & ";" & meNr & ";"& meInit & ";" & "xx99123sy#z"

                  
                   strEkspHeaderB = strEkspHeaderB & tsa_txt_152 &" Opt. ~ dage;"& tsa_txt_317 &" >> dd. ~ dage;Afholdt ~ dage;Afholdt u. løn ~ dage;Udbetalt ~ dage;"& tsa_txt_281 &" "& tsa_txt_280 &" ~ dage;"

                   strEkspHeaderB = strEkspHeaderB & "xx99123sy#z"
                   end if


               end if




   
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	call TimeOutVersion()
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\afstem_tot.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "medarbafexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
           

                

				'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
                if show <> 4 then

                 strEkspHeader = replace(strEkspHeader, "xx99123sy#z", vbcrlf)
	            ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)

				objF.WriteLine(strEkspHeader)
				objF.WriteLine(ekspTxt)
				else

                strEkspHeaderA = replace(strEkspHeaderA, "xx99123sy#z", vbcrlf)
	            ekspTxtA = replace(ekspTxtA, "xx99123sy#z", vbcrlf)
                
                strEkspHeaderB = replace(strEkspHeaderB, "xx99123sy#z", vbcrlf)
	            ekspTxtB = replace(ekspTxtB, "xx99123sy#z", vbcrlf)


                objF.WriteLine(strEkspHeaderA)
				objF.WriteLine(ekspTxtA)
                objF.WriteLine(strEkspHeaderB)
				objF.WriteLine(ekspTxtB)
                end if
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
	

end if 

 


 
 

   
    
   %>
    </div>
    

<%








    if media <> "print" AND media <> "export" then
    ptop = 102
    pleft = 1020 '(globalWdt + 70)
    pwdt = 200

call eksportogprint(ptop,pleft,pwdt)
%>




<form action="#" method="post">
<tr> 

    <td valign=top align=center>
   <input type=image src="../ill/export1.png" onclick="popUp('afstem_tot.asp?media=export&usemrn=<%=intMid %>&show=<%=show%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&yuse=<%=yuse %>',400,200,200,100)" />
    </td>
    <td class=lille><input id="Submit3" type="button" value="Eksportér til .csv >> " style="font-size:9px; width:130px;" onclick="popUp('afstem_tot.asp?media=export&usemrn=<%=intMid %>&show=<%=show%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&yuse=<%=yuse %>',400,200,200,100)" /></td>
</tr>
</form>


<form action="afstem_tot.asp?media=print&usemrn=<%=intMid %>&show=<%=show%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&yuse=<%=yuse %>" method="post" target="_blank">

<tr>

    <td valign=top align=center>
   <input type=image src="../ill/printer3.png"/>
    </td><td class=lille><input id="Submit6" type="submit" value="Print venlig" style="font-size:9px; width:130px;" /></td>
</tr>



</form>


   
	
   </table>
</div>

<%else 

 'Response.Write("<script language=""JavaScript"">window.print();</script>")

end if %>
	

    
 
   
  
   
<%end if 'validering %>
<!--#include file="../inc/regular/footer_inc.asp"-->






  