
<%

public fmLink, timertot, ucount
function ugeseddel(medarbid, stdatoSQL, sldatoSQL, visning, showtotal, showheader)

st_dato = day(stdatoSQL) &"-"& month(stdatoSQL) & "-"& year(stdatoSQL)

    stdatoSQL = year(stdatoSQL) &"-"& month(stdatoSQL) & "-" & day(stdatoSQL)
    sldatoSQL = year(sldatoSQL) &"-"& month(sldatoSQL) & "-" & day(sldatoSQL)




'** Altid i den rigtige uge uanset om det er slutdaot eller startato eller om ugen skifter år så den. f.eks starter i 2014 og slutter i 2015
varTjDatoUS_tor = dateAdd("d", 3, varTjDatoUS_man)

%>







 <!-- **** Ugeseddel *** -->

<%if showheader = 1 then %>
   <table cellpadding=0 cellspacing=0 border=0 width="100%">
   <tr><td style="padding:20px 10px 0px 0px;">
   
    <%
               'if thisfile = "ugeseddel_2011.asp" then
               fmLink = "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu 
               'else
               'fmLink = "weekpage_2010.asp?medarbid="&medarbid&"&st_dato="&st_dato
               'end if


    oimg = "blank.gif" '"ikon_stat_joblog_48.png"
	oleft = 0
	otop = 0
	owdt = 500
	oskrift = tsa_txt_338 &" "& datepart("ww", varTjDatoUS_tor, 2, 2) &" - "& datepart("yyyy", varTjDatoUS_tor, 2, 2) 
	
	'call sideoverskrift_2014(oleft, otop, owdt, oskrift)

    call meStamdata(usemrn)     
	
    %>
       <h4><%=oskrift &"<br><span style=""font-size:11px; font-weight:lighter;"">"& meTxt & "</span>"%></h4>

       <br />&nbsp;

    </td>




        <%
       thisWeek = datepart("ww", sldatoSQL, 2,2)
       prevWeek = datepart("ww", dateAdd("d", -7, sldatoSQL), 2,2) 
       nextWeek = datepart("ww", dateAdd("d", 7, sldatoSQL), 2,2) 
       %>
    
	<td align=right>
	
	

    <%if media <> "print" then %>
    
  
	<table cellpadding=0 cellspacing=0 border=0 width=180>
	<tr>
	<td align=right style="padding:0px 10px 0px 0px;"><a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&nomenu=<%=nomenu %>">< uge <%=prevWeek %></a></td>
        <td align="center" style="border:1px #000000 solid; padding:5px;"><b><%=thisWeek %></b></td>
   <td style="padding-right:10px;" align=right><a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&nomenu=<%=nomenu %>">uge <%=nextWeek %> ></a></td>
	</tr>
	</table>
    <%end if %>
	

  
	</td>
     </tr>


    <!-- afslut uge -->

       <%
       visLukUge = 0    
       if visLukUge = 1 then
       %>
       <tr><td colspan="2" style="padding-top:20px;">
   
       <%
  
          if cint(smilaktiv) <> 0 AND media <> "print" then

           varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man) 'sldatoSQL 'dateAdd("d", 7, sldatoSQL) 'FEJL ved årsskioft, detteaar bliver forkert

         call medrabSmilord(usemrn)
         call smiley_agg_fn()
         call smileyAfslutSettings()
         call meStamdata(usemrn)

          '** Henter timer i den uge der skal afsluttes ***'
          call afsluttedeUger(year(now), usemrn)

         '** Er kriterie for afslutuge mødt? Ifht. medatype mindstimer og der må ikke være herreløse timer fra. f.eks TT
         call timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, afslutugekri, usemrn)


         call timerDenneUge(usemrn, lto, tjkTimeriUgeSQL, akttypeKrism)
     
         'response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth

         if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
         weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
         else
         weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
         end if

        '** er perioden afsluttet

         call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2), weekMonthDate, usemrn) 


        call smileyAfslutBtn(SmiWeekOrMonth)

        call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf) 





         '**** Afslut uge ***'
	    '**** Smiley vises hvis sidste uge ikke er afsluttet, og dag for afslutninger ovwerskreddet ***' 12 
        if cint(SmiWeekOrMonth) = 0 then
        denneUgeDag = datePart("w", now, 2,2)
        s0Show_sidstedagsidsteuge = dateAdd("d", -denneUgeDag, now) 'now
        else
        s0Show_sidstedagsidsteuge = now
        end if

        '** finder kriterie for rettidig afslutning
        call rettidigafsl(s0Show_sidstedagsidsteuge)

        if cint(SmiWeekOrMonth) = 0 then
            s0Show_weekMd = datePart("ww", s0Show_sidstedagsidsteuge, 2,2)
        else
            s0Show_weekMd = datePart("m", s0Show_sidstedagsidsteuge, 2,2)
        end if

        
        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
		call erugeAfslutte(year(s0Show_sidstedagsidsteuge), s0Show_weekMd, usemrn)


       

     

       

        

     
            
	    '**** Afslut uge ***'
	    '**** Smiley vises om mandagen, første gang me logger på, hvis sidste uge ikke er afsluttet ***'
	    if (cDate(formatdatetime(now, 2)) >= cDate(formatdatetime(cDateUge, 2)) AND cint(ugeNrStatus) = 0) OR cint(smiley_agg) = 1 then
        'if (datepart("w", now, 2,2) = 1 AND datepart("h", now, 2,2) <= 23 AND session("smvist") <> "j") OR cint(smiley_agg) = 1 then
    	smVzb = "visible"
    	smDsp = ""
    	session("smvist") = "j"
    	showweekmsg = "j"
    	else
    	smVzb = "hidden"
    	smDsp = "none"
    	showweekmsg = "n"
    	end if
    	
    	

          
              

    	    
            '*** Auto popup ThhisWEEK now SMILEY
	      %>
	        <div id="s0" style="position:relative; left:20px; top:6px; width:725px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:2; background-color:#FFFFFF; padding:10px; border:0px #CCCCCC solid;">
	      
                 <%

               
           '*** Viser sidste uge
            'weekSelected = tjekdag(7)

            'if session("mid") = 1 then
            'response.write "varTjDatoUS_son " & varTjDatoUS_son & "<br>"
            'end if

            '*** Viser denne uge
            weekSelectedThis = dateAdd("d", 7, now) 

	        call showsmiley(weekSelectedThis, 1)

            call afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto)

            
               
            if cint(afslutugekri) = 0 OR ((cint(afslutugekri) = 1 OR cint(afslutugekri) = 2) AND cint(afslProcKri) = 1) OR cint(level) = 1 then 
            
                

	        call afslutuge(weekSelectedTjk, 1, varTjDatoUS_son, "ugeseddel")

             
         

            else
            %>

                  <%
             '** Status på antal registrederede projekttimer i den pågældende uge    
             select case lto
             case "tec", "esn"
             case else %>

                
          <%if cint(afslProcKri) = 1 then
                   sm_bdcol = "#DCF5BD"
                   else
                    sm_bdcol = "mistyrose"
                   end if %>



                <div style="color:#000000; background-color:<%=sm_bdcol%>; padding:10px; border:0px <%=sm_bdcol%> solid;">
            <%=tsa_txt_398 & ": "& totTimerWeek %> 
                
                 <%if afslutugekri = 2 then %>
                    <%=tsa_txt_399 %>
                    <%end if %>
                
                <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %> = <b><%=afslProc %> %</b> <%=tsa_txt_095 %> <b><%=datePart("ww", tjkTimeriUgeDtTxt, 2,2) %></b>
                <br />(<%=left(weekdayname(weekday(formatdatetime(tjkTimeriUgeDt, 2))), 3) &". "& formatdatetime(tjkTimeriUgeDt, 2)%> - <%= left(weekdayname(weekday(formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2))), 3) &". "&formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2) %>)
               
                 <%=tsa_txt_400 %>: <b><%=afslutugekri_proc %> %</b>  <%=tsa_txt_401%>.
               </div>
                <br />
            <%
                end select


            end if
	        
	        %>
               
	      <br /><br />
	        <span id="se_uegeafls_a" style="color:#5582d2;">[+] <%=tsa_txt_402 & " "& year(varTjDatoUS_son) %></span><br /><%

            varTjDatoUS_ons = dateAdd("d", -3, varTjDatoUS_son)

	        useYear = year(varTjDatoUS_ons)
            '** Hvilke uger er afsluttet '***
	        call smileystatus(usemrn, 1, useYear)
	        
                
                
                
             %>
              
	        <br />&nbsp;

               
	        </div>
        
        <br /><br />
        <%end if %>


        <% if cint(smilaktiv) <> 0 AND media = "print" then 

         call smileyAfslutSettings()

        if cint(SmiWeekOrMonth) = 0 then
        sm_sidsteugedag = datePart("ww",varTjDatoUS_son, 2,2)
        else
        sm_sidsteugedag = datePart("m",varTjDatoUS_son, 2,2)
        end if

             call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2) , sm_sidsteugedag, usemrn) 

             call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf) 
        
        end if %>
	</td>


      
  
	
    </tr>
    </table>
  <%end if 
      
   end if   'visLukUge%>

  

<%if media <>  "print" then '  AND cint(showheader) = 100 then 'slået fra
    
   
    %>

  <form action="ugeseddel_2011.asp?func=opdaterstatus&varTjDatoUS_man=<%=varTjDatoUS_man %>&usemrn=<%=usemrn %>&nomenu=<%=nomenu%>" method="post">
  
<%end if%>

 
  <table cellpadding=0 cellspacing=0 border=0 width=100%>
      

   <tr><td valign=top colspan=2>

         <%if showheader <> 1 then %>
       <br /><br /><br /><br />
       <%end if %>


       <b><%=weekdayname(weekday(st_dato, 1)) &" "& formatdatetime(st_dato, 1)%></b><br />


   
   <table width="100%" class="table dataTable table-striped table-hover ui-datatable">
   <tr >
            <th style="width:100px;"><b><%=tsa_txt_183 %></b></th>
            <th style="width:150px;"><b><%=tsa_txt_237 %></b></th>
            <th style="width:100px;"><b><%=tsa_txt_022 %> (<%=tsa_txt_066 %>)</b></th>
            <th style="width:150px;"><b><%=tsa_txt_244 %></b></th>
            <th style="width:100px;"><b><%=tsa_txt_068 %></b> <%=tsa_txt_329 %><br /><%=tsa_txt_296 %></th>
            <th style="width:100px; text-align:right;"><b><%=tsa_txt_148 %></b></th>

       <%'if showheader = 1 then 'slået fra %>
            <td class=alt align="right" style="width:60px; padding-left:20px;">

             
                
                
                <%if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then%>

                
                <label>Status <input type="checkbox" class="status_all" id="status_all_<%=st_dato%>" value="1" /></label><br />

                

                <select id="status_sel" name="FM_godkendt" style="width:60px;">
                    <option value="0">Ingen</option>
                    <option value="1">Godkendt</option>
                    <option value="2">Afvist</option>
                </select>
                <%else %>
                Status
                <%end if %>
            </td>
       <!--else %>
             <td class=alt style="width:60px; padding-left:20px;">&nbsp;</td>
       <end if %>-->
   </tr>

 

    
   <%'*** Ugeseddel ***' 
   
   call afsluger(medarbid, stdatoSQL, sldatoSQL)

    select case lto
    case "nonstop"
    strSQLodrBy = " tdato, tfaktim, sttid "
    case else
    strSQLodrBy = " tdato, sttid "
    end select


   strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
   &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, godkendtstatusaf, m.mnr, timerkom, init, a.fase, a.easyreg, j.risiko, sttid, sltid FROM timer "_
   &" LEFT JOIN medarbejdere AS m ON (m.mid = tmnr)"_
   &" LEFT JOIN aktiviteter AS a ON (a.id = taktivitetid)"_
   &" LEFT JOIN job AS j ON (j.jobnr = tjobnr)"_
   &" LEFT JOIN kunder K on (kid = tknr) WHERE tmnr = "& medarbid &" AND "_
   &" tdato BETWEEN '"& stdatoSQL &"' AND '"& sldatoSQL &"' ORDER BY "& strSQLodrBy &" DESC "
   
   'Response.Write strSQL
   'Response.flush
   at = 0
   if showheader = 1 then
   timertot = 0
   end if
   lwedaynm = ""
   timerDag = 0
   antalEasyReg = 0
   eaDag = 0
   atDag = 0
   'ugegodkendtTxt = ""
   
   oRec.open strSQL, oConn, 3
   while not oRec.EOF 


                '** Er periode/timeregistrering godkendt ***'
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

                
                call erugeAfslutte(useYear, usePeriod, usemrn)
                
                call lonKorsel_lukketPer(tjkDag, oRec("risiko"))
                
		       ' if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
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







    
    select case right(at, 1)
    case 0,2,4,6,8
    bgcol = "#ffffff"
    case else
    bgcol = "#EFF3ff"
    end select
    
    call akttyper2009prop(oRec("tfaktim"))
    
    
    if lwedaynm <> weekdayname(weekday(oRec("Tdato"))) then
     if at <> 0 then%>
    <tr>
        <td colspan=6 align=right><b><%=formatnumber(timerDag, 2) %> t.</b>
            <!--<br />
        <span style="color:#999999;">Fordelt på <%=atDag %> registreringer, heraf <%=eaDag %> Easyreg.</span><br /><br />&nbsp;--></td>
        <td>&nbsp;</td>

    </tr>
    <%
    atDag = 0
    eaDag = 0
    timerDag = 0
    
    end if 
        
     if visning <> 1 then%>
    <tr bgcolor="#D6Dff5">
        <td colspan=7 valign=bottom style="height:30px; border-bottom:1px #8CAAE6 solid; padding:3px;"><b><%=weekdayname(weekday(oRec("Tdato"), 1)) %></b></td>
    </tr>
    <%end if


    end if
    
    %>
   <tr bgcolor="<%=bgcol%>">
   <td valign=top style="border-bottom:1px #cccccc solid; padding-right:20px; height:20px; white-space:nowrap;">
       <%
        tmp = 1
       if tmp = 900 then
       'if (((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR (level = 1)) AND media <> "print") then %>
       <input type="text" value="<%=oRec("tdato")%>" style="width:80px;" />
        <%else %>
       <%=oRec("tdato")%>
       <%end if %>

   </td>
   <td valign=top style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=left(oRec("tmnavn"), 10) %> 
   <%if len(trim(oRec("init"))) <> 0 then %>
   [<%=oRec("init") %>]
   <%end if %></td>
   <td valign=top style="border-bottom:1px #cccccc solid; padding-right:20px;"><b><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</b></td>
   <td valign=top style="border-bottom:1px #cccccc solid; padding-right:20px;">
   <%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>

   <td valign=top style="border-bottom:1px #cccccc solid;"><%=oRec("taktivitetnavn") %>
   
   <%if oRec("easyreg") <> 0 then %>
    <span style="color:#999999;">(E!)</span>
   <%antalEasyReg = antalEasyReg + 1  
   eaDag = eaDag + 1%>
   <%end if %>
   
    
    <%
	call akttyper(oRec("tfaktim"), 1)
	%>

    <span style="color:yellowgreen;">(<%=akttypenavn%>)</span>
	
   
   <%if len(trim(oRec("fase"))) <> 0 then %>
   <br />
   <span style="color:#5582d2;"><%=replace(oRec("fase"), "_", " ")%></span>
   <%end if %>

    <%if len(trim(oRec("timerkom"))) <> 0 then %>
    <span style="color:#999999;">
    <%if len(oRec("timerkom")) > 100 then %>
    <i><%=left(oRec("timerkom"), 100) %>...</i>
    <%else %>
    <i><%=oRec("timerkom")%></i>
    <%end if %>
    </span>
    <%end if %>
   </td>
   <td align=right valign=top style="border-bottom:1px #cccccc solid;">
          <%
      
                
                
                if (((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR (level = 1)) AND media <> "print") then%>
                 <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=650,height=600,resizable=yes,scrollbars=yes')" class=vmenu><%=formatnumber(oRec("timer"), 2) %></a>
	            
                <%else %>
                 <%=formatnumber(oRec("timer"), 2) %>
                <%end if %>

                
                <%if formatdatetime(oRec("sttid"), 3) <> "00:00:00" then %>

                <br /><span style="color:#999999; font-size:9px;"><%=left(formatdatetime(oRec("sttid"), 3), 5) %> - <%=left(formatdatetime(oRec("sltid"), 3), 5) %></span>

                <%end if %>
            

                </td>
                <td align="right" valign="top" style="border-bottom:1px #cccccc solid;">
                    <%
                   
                'if cint(ugegodkendt) <> 1 AND media <> "print" then 
                
                '              erGk = 0
                '             gkCHK0 = ""
                '             gkCHK1 = ""
                '             gkCHK2 = ""
                '             gk2bgcol = "#999999"
                '             gk1bgcol = "#999999"
                '             gk0bgcol = "#999999"

                '         select case cint(oRec("godkendtstatus"))
                '         case 2
                '         gkCHK2 = "CHECKED"
                '         erGk = 2
                '         gk2bgcol = "red"
                '         case 1
                '         gkCHK1 = "CHECKED"
                '         erGk = 1
                '         gk1bgcol = "green"
                '         case else
                '         gkCHK0 = "CHECKED"
                '         erGk = 0
                '         gk0bgcol = "#000000"
                '         end select
                 
                '         erGkaf = ""


                
                %>

                    <!--

                   <input name="ids" id="ids" value="<%=oRec("tid")%>" type="hidden" />
					   <span style="color:<%=gk1bgcol%>;">Godk:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt1_<%=v%>" class="FM_godkendt_1" value="1" <%=gkCHK1 %>><br />
                       <span style="color:<%=gk2bgcol%>;">Afvist:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt2_<%=v%>" class="FM_godkendt_2" value="2" <%=gkCHK2 %>><br />
                       <span style="color:<%=gk0bgcol%>;">Ingen:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt0_<%=v%>" value="0" class="FM_godkendt_0" <%=gkCHK0 %>><br />

                    -->



                    <label>

                       <%

                           
                    
                       
                     '  else 
                       
                             'response.write "teamLederEditOK: "& teamLederEditOK & "ugeerAfsl_og_autogk_smil: " & ugeerAfsl_og_autogk_smil & " ugegodkendt: "& ugegodkendt &" level "& level 


                '***  Admin kan godkende for alle medarb. Teamleder kun for dem man er teamleder for. 
                 '*** visAlleMedarb = 1 skal være slået til før man kan godkende andres (undt. admin), da det ikke giver mening at goende sine egne timer
                 select case cint(oRec("godkendtstatus"))
                 case 2
                 gkTxt = "<b>!</b>"
                 gkbgcol = "red"
                     if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                     gkInputype = "CHECKBOX"
                     gkInputName = "ids"
                     else
                     gkInputype = "hidden"
                     gkInputName = "ids_hd"
                     end if
                 case 1
                 gkTxt = "<b><i>V</i></b>"
                 gkbgcol = "green"
                  if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                     gkInputype = "CHECKBOX"
                           gkInputName = "ids"
                     else
                     gkInputype = "hidden"
                           gkInputName = "ids_hd"
                     end if
                 case else
                 gkTxt = ""
                     if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                     gkInputype = "CHECKBOX"
                           gkInputName = "ids"
                     else
                     gkInputype = "hidden"
                           gkInputName = "ids_hd"
                     end if
                 gkbgcol = "#000000"
                 end select

                
                 
                 Response.write "<span style=""color:"& gkbgcol &";"">"& gkTxt &"</span>" 

                 select case cint(oRec("godkendtstatus"))
                 case 1, 2
                 %>
                 <br><span style='font-size:9px; color:#999999;'><%=left(oRec("godkendtstatusaf"), 12)  %></span><br />
                 <%
                  end select

                       
                       %>

                    <input class="status_chk_<%=st_dato%>" name="<%=gkInputName %>" type="<%=gkInputype %>" value="<%=oRec("tid")%>" /> 
                        </label>



                </td>
   </tr>
   
   <%
   lwedaynm = weekdayname(weekday(oRec("Tdato")))
   lwedate = oRec("Tdato")

   if cint(aty_real) = 1 then
   timertot = timertot + oRec("timer")
   timerDag = timerDag + oRec("timer")
   end if
   at = at + 1

   atDag = atDag + 1
   oRec.movenext
   wend
   oRec.close


       


   

   
   if at <> 0 then

           %>

             <tr>
                 <td colspan=5>&nbsp;</td>
        <td align=right><b><%=formatnumber(timerDag, 2) %> t.</b>
       <!-- <br />
        <span style="color:#999999;">Fordelt på <%=atDag %> registreringer, heraf <%=eaDag %> Easyreg.</span>--></td>
        <td>&nbsp;</td>
    </tr>

       <% end if
           
  
           
        if cint(showtotal) = 1 then %>
            <tr>
                <td colspan=5>&nbsp;</td>
                <td><br /><br />Total denne uge:<br /></td>
               <td>&nbsp;</td>
       
           </tr>
           <tr>
                <td colspan=5>&nbsp;</td>
                <td align=right style="background-color:yellowgreen;"><b><%=formatnumber(timertot, 2) %> t.</b></td>
               <td>&nbsp;</td>
       
           </tr>

          <%end if 'total%>

     
                <%if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then%>
           <tr>

           <td colspan="7" align="right" style="padding-top:40px;"><input type="submit" value="Opdater >>"</td>

       </tr>
        </form>
               <%end if %>


    

      

       <%
      


       if cint(visning) = 1 then


    '**** Timer uden match, fra import Eller timeTag **** sttid, sltid
     strSQL = "SELECT timer, jobid, tip.jobnavn, tip.aktnavn AS tipaktnavn, aktid, tdato, mnavn, mnr, init, timerkom, a.navn AS aktnavn, tip.id AS extsysId, errmsg "_
     &" FROM timer_import_temp AS tip "_
     &" LEFT JOIN medarbejdere AS m ON (m.mid = tip.medarbejderid)"_
     &" LEFT JOIN job AS j ON (j.id = tip.jobid)"_
     &" LEFT JOIN aktiviteter AS a ON (a.id = tip.aktid)"_
     &" WHERE tdato = '"& stdatoSQL & "' AND medarbejderid = "& medarbid & " AND overfort = 0"
     

    'if (lto = "assurator" AND session("mid") = 14) then
    'response.write " strSQL: " & strSQL
    ' response.flush
    'end if
           
    if len(trim(ucount)) <> 0 then
       u = ucount
     else
       u = 0
    end if

        
           
        oRec.open strSQL, oConn, 3
        while not oRec.EOF 

       select case right(u, 1)
       case 0,2,4,6,8
       timerTempBgCol = "#FFE6E6" '"#FFDFFF"
       case else
       timerTempBgCol = "#FFFFFF"
       end select


      
        
     

           if media <> "print" then%>
               <form action="timereg_akt_2006.asp?func=db&rdir=ugeseddel_2011&usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post">
               <input type="hidden" value="<%=positiv_aktivering_akt_val%>" name="FM_pa" />
               <input type="hidden" id="Hidden4" name="FM_dager" value=", xx"/>
               <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
               <input type="hidden" id="Hidden5" name="year" value="<%=year(oRec("tdato")) %>"/>
              
            <%end if%>

                <%if lastDato <> oRec("tdato") then %>
               <tr>

                   <td colspan="7"><span style="color:red;">Registreringer uden match:</span></td>

               </tr>
                    <%end if %>
      
       
     
       <tr style="background-color:<%=timerTempBgCol%>;">
           <td valign="top" style="border-bottom:1px #cccccc solid;">
                <input type="hidden" id="extsysId" name="extsysId" value="<%=oRec("extsysId") %>"/>
               <input type="text" name="FM_datoer" value="<%=oRec("tdato")%>" style="width:80px;" /></td>
           <td valign="top" style="border-bottom:1px #cccccc solid;"><input type="hidden" value="<%=usemrn %>" name="FM_medid" id="FM_medid" /><%=left(oRec("mnavn"), 10) %>
                <%if len(trim(oRec("init"))) <> 0 then %>
   [<%=oRec("init") %>]
   <%end if %>

           </td>
           

           <%if len(trim(oRec("jobid"))) <> 0 AND oRec("jobid") <> 0 then
               ugsJobid = oRec("jobid") 
               ugsJobidFundet = 0

                strSQLjobf = "SELECT id FROM job WHERE id = "& ugsJobid
                oRec6.open strSQLjobf, oConn, 3
                if not oRec6.EOF then 

                   ugsJobidFundet = 1

                end if
                oRec6.close

             else
               ugsJobid = 0
               end if %>

           <td valign="top" colspan="2" style="border-bottom:1px #cccccc solid; white-space:nowrap;">
               <input type="hidden" name="FM_jobid" id="FM_jobid_<%=u %>" value="<%=ugsJobid %>"/>
               <input type="hidden" name="FM_job_opr" id="FM_job_opr_<%=u %>" value="<%=oRec("jobnavn") %>"/>
               <input type="text" name="FM_jobnavn" id="FM_job_<%=u %>" class="FM_job" value="<%=oRec("jobnavn") %>" style="width:250px;" />
               <%if cdbl(ugsJobid) <> 0 AND cint(ugsJobidFundet) = 1 then %>
               <span style="font-size:16px; color:yellowgreen"><i>V</i></span>
               <%end if %>

               <div id="dv_job_<%=u %>" class="dv_job" style="position:absolute; background-color:#FFFFFF; visibility:hidden; display:none; width:250px; height:200px; overflow:auto; padding:10px; border:1px #999999 solid;"></div>
               <br /><span style="color:#999999; font-size:9px;"><b>Err. msg: </b><br /><%=oRec("errmsg") %></span>
           </td>
           <td valign="top" style="border-bottom:1px #cccccc solid; white-space:nowrap;">
           <!--<input type="hidden" name="FM_aktivitetid" id="FM_aktid_<%=u %>" value="0"/>-->
    
            <%if len(trim(oRec("aktnavn"))) <> 0 AND oRec("aktnavn") <> "" then
               ugeAktnavn = oRec("aktnavn") 
             else
               ugeAktnavn = oRec("tipaktnavn")
             end if %>


                <%if len(trim(oRec("aktid"))) <> 0 AND oRec("aktid") <> 0 then
               ugsAktid = oRec("aktid") 


                     ugsAktidFundet = 0

                strSQLaf = "SELECT id FROM aktiviteter WHERE id = "& ugsAktid
                oRec6.open strSQLaf, oConn, 3
                if not oRec6.EOF then 

                   ugsAktidFundet = 1

                end if
                oRec6.close

               else
               ugsAktid = 0
               end if %>
            
               
           <input type="hidden" name="FM_aktivitetid" id="FM_aktid_<%=u %>" value="<%=ugsAktid %>"/>
           <input type="text" name="FM_aktnavn" value="<%=ugeAktnavn %>" class="FM_akt" id="FM_akt_<%=u %>" style="width:200px;" />

                <%if cdbl(ugsAktid) <> 0 AND cint(ugsAktidFundet) = 1 then %>
               <span style="font-size:16px; color:yellowgreen"><i>V</i></span>
               <%end if %>
              
               <div id="dv_akt_<%=u %>" class="dv_akt" style="position:absolute; background-color:#FFFFFF; visibility:hidden; display:none; width:200px; height:200px; overflow:auto; padding:10px; border:1px #999999 solid;"></div>
                 
               <br /><img src="../ill/blank.gif" width="1" height="5" /><br />
               <textarea id="FM_kom" name="FM_kom_0" style="width:200px; height:40px;" placeholder="Kommentar"><%=trim(oRec("timerkom")) %></textarea>
               

            
           </td>
           <td valign="top" style="border-bottom:1px #cccccc solid;" align="right"><input type="text" name="FM_timer" value="<%=formatnumber(oRec("timer"), 2) %>" style="width:60px;" />
               <input type="hidden" name="FM_timer" value="xx"/>
           </td>
           <td valign="top" style="border-bottom:1px #cccccc solid;" align="right"><input type="submit" value="Opdater >>" style="font-size:9px;" />
             
               <%if media <> "print" then %>  
                 <br /><br /><br />
               <a href="ugeseddel_2011.asp?func=slet_tip&usemrn=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man%>&varTjDatoUS_son=<%=varTjDatoUS_son %>&nomenu=<%=nomenu%>&id=<%=oRec("extsysId") %>" class="red">X</a>
               <%end if %>
            
           </td>

       </tr>
           
       <%if media <> "print" then %>   
    
       </form>
   
       <% end if
        lastDato = oRec("tdato") 
        timerUge = timerUge + timerDag 
        u = u + 1
        ucount = u

   oRec.movenext
   wend
   oRec.close


   end if

    %>



         </table>
      
      
       


   </td>


  

  
   
   </tr>
    </table>
    
   
    

    <%end function

%>