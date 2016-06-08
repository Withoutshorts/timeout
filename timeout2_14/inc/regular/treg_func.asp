<%
function tregsubmenu()

%>

<a href="timereg_akt_2006.asp?showakt=1&hideallbut_first=2" class=rmenu><%=tsa_txt_116 %></a>
	<%if level <= 7 then %>
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="materialer_indtast.asp?id=0&fromsdsk=0&aftid=0" class=rmenu target="_blank"><%=tsa_txt_191 %></a>
	<%end if %>
	
	<!--
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="timereg_akt_2006.asp?showakt=0" class=rmenu target="a"><%=tsa_txt_117 %></a>
	-->
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="joblog_timetotaler.asp?nomenu=1" class=rmenu target="_blank"><%=tsa_txt_119 %></a>
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="joblog.asp?nomenu=1" class=rmenu target="_blank"><%=tsa_txt_118 %></a>
	
	
	    <%if level = 1 then %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='feriekalender.asp?menu=job' class='rmenu' target="_top">Ferie, Afspad. & Sygdom's kalender</a>
		<%else %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='feriekalender.asp?menu=job' class='rmenu' target="_top">Ferie & Afspad. kalender</a>
		<%end if %>

        <%if level <= 2 OR level = 6 then %>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="week_real_norm_2010.asp" class=rmenu target="_top"><%=tsa_txt_356 %></a>
	    <%end if %>

          <%if level <= 7 then %>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="ressource_belaeg_jbpla.asp?menu=webblik" class=rmenu target="_top">Ressource Forecast (timebudget)</a>
	    <%end if %>
       
     
<%

jobasnvigv = 0
strSQLigv = "SELECT jobasnvigv FROM licens WHERE id = 1"
oRec.open strSQLigv, oConn, 3
if not oRec.EOF then
jobasnvigv = oRec("jobasnvigv")
end if
oRec.close

if jobasnvigv = 1 then %>
&nbsp;&nbsp;|&nbsp;&nbsp;<a href="stat_opdater_igv.asp?func=opdater" class='rmenu' target="_blank">Stade-indmeldinger</a>
<%end if 

      call meStamdata(session("mid")) 
      if cint(meVisskiftversion) = 1 then

    %>
 &nbsp;&nbsp;&nbsp;&nbsp;<span style="background-color:#F7F7F7; padding:2px;">Skift: 
<%
        
       if lto = "biofac" then
     
        %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=support@outzource.dk&key=2.2013-0912-TO146" class=rmenu target="_top">Adminbruger </a>
       &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=me@outzource.dk&key=2.2013-0912-TO146" class=rmenu target="_top">Medarbejder </a>
        <%
       end if

        if lto = "demo" then
        %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=support@outzource.dk&key=2.052-xxxx-B000" class=rmenu target="_top">Adminbruger </a>
       &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=me@outzource.dk&key=2.052-xxxx-B000" class=rmenu target="_top">Medarbejder </a>
        <%
       end if


        if lto = "outz" OR lto = "intranet - local" then
       
      %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" class=rmenu target="_top">EPI OSL</a>
       <%
       end if

       if lto = "epi" then
     
       %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1010-TO135" class=rmenu target="_top">EPI STA</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" class=rmenu target="_top">EPI OSL</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" class=rmenu target="_top">EPI AB</a>
      
       <%
     
       end if

        if lto = "epi_no" then
     
      %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1010-TO135" class=rmenu target="_top">EPI STA</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" class=rmenu target="_top">EPI</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" class=rmenu target="_top">EPI AB</a>
      
       <%
   
       end if

       if lto = "epi_ab" then
      
       %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1010-TO135" class=rmenu target="_top">EPI STA</a>
       &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" class=rmenu target="_top">EPI</a>
       &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" class=rmenu target="_top">EPI OSL</a>
       <%
       
       end if

        if lto = "epi_sta" then
     
    %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" class=rmenu target="_top">EPI OSL</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" class=rmenu target="_top">EPI</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;<a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" class=rmenu target="_top">EPI AB</a>
      
       <%
       
       end if


       if lto = "mmmi" then
      
      %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2011-0410-TO127" class=rmenu target="_top">UNIK</a>
       <%end if
       


       if lto = "unik" then
    
      %>
       <a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-0311-TO119" class=rmenu target="_top">MMMI</a>
       <%
       end if
       
           %>&nbsp;</span><%


    end if

end function

public lnkUge
function treg_3menu(thisfile)

    lnkTreg = "timereg_akt_2006.asp?showakt=1&hideallbut_first=2&strdag="&day(varTjDatoUS_man)&"&strmrd="& month(varTjDatoUS_man) &"&straar="&year(varTjDatoUS_man)
    lnkUge = "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son
	lnkAfstem = "afstem_tot.asp?usemrn="&usemrn&"&show=5&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son
    lnkLogind = "logindhist_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man &"&varTjDatoUS_son="&varTjDatoUS_son

select case thisfile
case "timereg_akt_2006"

 
  
	bgtreg = "#ffff99"
    bgUge = "#FFFFFF"
    bgafst = "#FFFFFF"
	bgstemp = "#FFFFFF"
    toppx = 120
   

case "ugeseddel_2011.asp"

    
	bgtreg = "#ffffff"
    bgUge = "#FFFF99"
    bgafst = "#FFFFFF"
	bgstemp = "#FFFFFF"
    toppx = 120
  
   
case "afstem_tot.asp"

    bgtreg = "#ffffff"
    bgUge = "#FFFFff"
    bgafst = "#FFFF99"
	bgstemp = "#FFFFFF"
    toppx = 120


case "logindhist_2011.asp"  
    

     bgtreg = "#ffffff"
    bgUge = "#FFFFff"
    bgafst = "#FFFFFF"
	bgstemp = "#FFFF99"
    toppx = 120

end select

%>
<div id=tregmenu style="position:relative; left:21px; top:<%=toppx%>px;">
    <!--showakt:<%=showakt %>-->
    <table cellpadding=0 cellspacing=0 border=0>
        <tr>
            <td align=center id="treg" style="white-space:nowrap; width:100px; padding:2px; background-color:<%=bgtreg%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="<%=lnkTreg %>" id="showtim" class=rmenu><%=tsa_txt_116 %></a></td>
             
             <!--
             <td align=center id="afst" style="white-space:nowrap; width:100px; padding:4px; background-color:<%=bgafst%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="showaf" class=vmenu><%=tsa_txt_337 %></a></td>
			-->

             <td align=center id="Td2" style="white-space:nowrap; width:100px; padding:2px; background-color:<%=bgUge%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="<%=lnkUge %>" id="A1" class=rmenu><%=tsa_txt_337 %></a></td>
			

			 <td align=center id="Td1" style="white-space:nowrap; width:140px; padding:2px; background-color:<%=bgafst%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a id="afstemtot" href="<%=lnkAfstem %>" class=rmenu><%=tsa_txt_389 %></a>
			</td>
			<%
            if session("stempelur") <> 0 AND session("dontLogind") <> 1 then %>

              <!--
             <td align=center id="stemp" style="white-space:nowrap; width:100px; padding:4px; background-color:<%=bgstemp%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="showst" class=vmenu><%=tsa_txt_340 %></a></td>
            -->

              <td align=center id="td3" style="white-space:nowrap; width:100px; padding:2px; background-color:<%=bgstemp%>; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="<%=lnkLogind %>" id="a2" class=rmenu><%=tsa_txt_340 %></a></td>
			
			<%end if %>
        </tr>
    </table>
    </div>

<%end function %>