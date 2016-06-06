<%response.Buffer = true 
tloadA = now
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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
	thisfile = "logindhist_2011.asp"
    rdir = "lgnhist"
	media = request("media")

    'usemrn = request("usemrn")
    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if

    level = session("rettigheder")
 

    

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
                    'usemrn = session("mid")
                    end if


    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb


    call medarb_teamlederfor


    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then
    nomenu = 1
    else
    nomenu = 0
    end if

   

    varTjDatoUS_man = request("varTjDatoUS_man")
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'varTjDatoUS_son = request("varTjDatoUS_son")

   

    lnk = "&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&rdir="&rdir&"&nomenu="&nomenu


    'Response.Write varTjDatoUS_man
    'Response.end

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)


	'Response.Write "media " & media
	
    d_end = 6
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	
    select case func 
	case "-"
	
	case "xxxopdaterlist"

    Response.write Request("logindhist_id") & "<br>"
    Response.write Request("logindhist_dato") & "<br>"
    Response.Write Request("logindhist_sttid")  & "<br>"
    Response.Write Request("logindhist_sltid")  & "<br>"

    ids = split(Request("logindhist_id"), ", ")
    datoers = split(Request("logindhist_dato"), ", ")
    stTiders = split(Request("logindhist_sttid"), ", ")
    slTiders = split(Request("logindhist_sltid"), ", ")

    for i = 0 TO UBOUND(ids)

    '** start tid **'
    dato_tidSt = datoers(i) &" "& stTiders(i)
    if len(trim(stTiders(i))) <> 0 then
        call logindStatus(usemrn, 1, 2, dato_tidSt)
    end if

    Response.Flush
    
     '** slut tid **'
    dato_tidSl = datoers(i) &" "& slTiders(i)
    if len(trim(slTiders(i))) <> 0 then
        call logindStatus(usemrn, 1, 2, dato_tidSl)
    end if
    
    next


    Response.end
	

    case else
       

    call akttyper2009(2)
	
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
    if nomenu <> 1 then

	leftPos = 90
	topPos = 102

    else

    leftPos = 20
	topPos = 48

    end if
	
	%>

	

     <%call browsertype()
    
     %>

	     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <SCRIPT language=javascript src="inc/smiley_jav.js"></script>
    <SCRIPT language=javascript src="inc/logindhist_2011_jav.js"></script>
  

  <%
      if nomenu <> 1 then

    if media <> "print" then
    %>

     <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>

	</div>

    <% end if
	

	
         call menu_2014() 


      end if

    
       
    else 
	
	leftPos = 0
	topPos = 0
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
    <SCRIPT language=javascript src="inc/logindhist_2011_jav.js"></script>
	<%end if%>
	
	<%end if%>
	

	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">

        
          <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; width:75px; top:-20px; left:730px; z-index:0;"><a href="<%=lnkTimeregside%>" class="vmenu"><%=left(tsa_txt_031, 7) %>.</a></div>
    
  
    <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; top:-20px; width:85px; left:820px; z-index:0;"><a href="<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a></div>
  


  <%
    

	
	call ersmileyaktiv()

    if media <> "print" then
    tTop = 0
	tLeft = 0
    else
    tTop = 0
	tLeft = 0
    end if
	
    tWdth = 900
	
	
	call tableDiv(tTop,tLeft,tWdth)

    
      

    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16
      %>
    
	<form id="filterkri" method="post" action="logindhist_2011.asp">
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
      
	
	<%end if 'media
        %>

   
  

    <table cellpadding=0 cellspacing=0 border=0 width=100%>
     
    <%if media <> "print" then %>

       <%
       prevWeek = datepart("ww", prev_varTjDatoUS_man, 2,2) 
       nextWeek = datepart("ww", next_varTjDatoUS_man, 2,2) 
       %>

         <tr>
          <td><h3>Komme / Gå tid</h3></td>
          <td align=right valign=top style="padding:10px 10px 0px 0px;">
	    <table cellpadding=0 cellspacing=0 border=0 width=120>
	    <tr>
	    <td valign=top align=right><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man%>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>">< uge <%=prevWeek %></a></td>
       <td style="padding-left:20px;" valign=top align=right><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>">uge <%=nextWeek %> ></a></td>
	    </tr>
	    </table>

              </td>

      </tr>
        <%if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then %>
        <tr>
            <td valign=top style="padding:10px 10px 0px 0px;">
                   
                     <div style="width:600px;"><%
                    
                    if cint(smilaktiv) <> 0 AND media <> "print" then
                    call afslutMsgTxt 
            end if %></div>
                <br />&nbsp;
            </td>

        </tr>
        <%end if %>

    <%end if %>
	
        
          
    <tr><td colspan="2" style="padding-top:10px;">
        <%
            call smileyAfslutSettings()
            'Skal det være mandag for måned og søndag for uge??
            if cint(SmiWeekOrMonth) = 0 then
            periodeTxt = tsa_txt_005 & " "& datepart("ww", tjkDato, 2 ,2)
            weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
            else
            periodeTxt = monthname(datepart("m", tjkDato, 2 ,2))
            weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
            end if




        if cint(smilaktiv) <> 0 AND media <> "print" then

         call medrabSmilord(usemrn)
         call smiley_agg_fn()
         call meStamdata(usemrn)

          '** Henter timer i den uge der skal afsluttes ***'
          call afsluttedeUger(year(now), usemrn)

        '** Er kriterie for afslutuge mødt? Ifht. medatype mindstimer og der må ikke være herreløse timer fra. f.eks TT
         call timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, afslutugekri, usemrn)


         call timerDenneUge(usemrn, lto, tjkTimeriUgeSQL, akttypeKrism)
     
       
         if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
         weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
         else
         weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
         end if

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
	        <div id="s0" style="position:relative; left:0px; top:5px; width:725px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:2; background-color:#FFFFFF; padding:10px; border:0px #CCCCCC solid;">
	       <%
           '*** Viser sidste uge
            'weekSelected = tjekdag(7)

            '*** Viser denne uge
            weekSelectedThis = dateAdd("d", 7, now) 

	        call showsmiley(weekSelectedThis, 1)

            call afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto)

            
               
            if cint(afslutugekri) = 0 OR ((cint(afslutugekri) = 1 OR cint(afslutugekri) = 2) AND cint(afslProcKri) = 1) OR cint(level) = 1 then 
            
                

	        call afslutuge(weekSelectedTjk, 1, varTjDatoUS_son, "logindhist")

             
         

            else
           


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
	        <span id="se_uegeafls_a" style="color:#5582d2;">[+] <%=tsa_txt_402 & " "& year(varTjDatoUS_son)%></span><br /><%

            varTjDatoUS_ons = dateAdd("d", -3, varTjDatoUS_son)

	        useYear = year(varTjDatoUS_ons)
            '** Hvilke uger er afsluttet '***
	        call smileystatus(usemrn, 1, useYear)
	        
                
                
                
             %>
              
	        <br />&nbsp;

               
	        </div>
        
        <br /><br />
        <%end if %>

        <%if cint(smilaktiv) <> 0 AND media = "print" then 

           
          call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2) , weekMonthDate, usemrn) 

          call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf) 

           

        end if %>
         <br /><br />
        &nbsp;
        </td></tr>

    
	<tr>
	<td valign=top colspan=2 style="padding:10px;">
	
	
	<%call stempelurlist(useMrn, 0, 1, varTjDatoUS_man, varTjDatoUS_son, 0, d_end, lnk) %>
	
	</td></tr>

         <tr>
     <td valign=top colspan=2 style="padding:10px;">
     
	<%call fLonTimerPer(varTjDatoUS_man, 7, 0, useMrn) %>
	
	
	
	</td>
	</tr>
     

    </table>
    
 
 
    
  

   <%
	
	
	
	 if media <> "print" then

ptop = 0
pleft = 945
pwdt = 220

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&varTjDatoUS_son=<%=varTjDatoUS_son %>&media=print" method="post" target="_blank">
<tr> 
    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value=" Print venlig >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

        <tr>
            <td colspan="2" style="padding-top:40px;">

                 <%
    
        if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 then 

        sm_aar = year(varTjDatoUS_son)
        sm_mid = usemrn

        'response.write "weekMonthDate: "& weekMonthDate

        call erugeAfslutte(sm_aar, weekMonthDate, sm_mid)
                     
        fmlink = "ugeseddel_2011.asp?usemrn="& usemrn &"&varTjDatoUS_man="& varTjDatoUS_man &"&varTjDatoUS_son= "& varTjDatoUS_son &"&nomenu="& nomenu &"&rdir=logindhist"
        %>
           
        <%call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
        
       <%end if 'level %>
            </td>

        </tr>
   
	
   </table>
</div>
<%else

Response.Write("<script language=""JavaScript"">window.print();</script>")

end if %>
  

    </div><!-- table div>-->
    <br /><br />&nbsp;
    </div><!-- side div>-->
    <br /><br />&nbsp;
    <%
	
	
	end select
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
