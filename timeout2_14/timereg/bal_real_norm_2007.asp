<%response.Buffer = true 
tloadA = now
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
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
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "bal_real_norm_2007.asp"
	media = request("media")
	

  

	'Response.Write "media " & media
	
	call erStempelurOn()
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	sendemail = request("sendemail")
	
	select case func 
	case "-"
	
	case "lukper", "lukper_ok"


   
	
	lukdag = request("FM_lk_dag") 
	lukmd = request("FM_lk_md")
	lukaar = request("FM_lk_aar")

    if len(trim(request("FM_nulstil"))) <> 0 then
    nulstil = 1
    else
    nulstil = 0
    end if

    '** Er det første loop **'
    if request("lmt") <> "" then
    lmt = request("lmt")
    else
    lmt = 0
    end if

    if len(request("lastMid")) <> 0 then
    lastMid = request("lastMid")
    else
    lastMid = 0
    end if

 
	
    nextLmt = lmt + 1


	lukDato = lukaar & "/" & lukmd & "/" & lukdag

    lukDato_1 = dateAdd("d", 1, lukDato)
	
    if func = "lukper_ok" then
    gkTxt = "<span style=""color:yellowgreen;""><i>Ok</i></span>"
    else
    gkTxt = ""
    end if

	if isDate(lukDato) then
	 
	

     '********************************************************'
     '***** Overfører saldo (Korrektion til ny periode) ******'
     '********************************************************'

      '*** Finder startDato for afstemning


             call licensStartDato()

           

             call akttyper2009(6)
	         akttype_sel = "#-99#, " & aktiveTyper
	         akttype_sel_len = len(akttype_sel)
	         left_akttype_sel = left(akttype_sel, akttype_sel_len - 2)
	         akttype_sel = left_akttype_sel


             if func = "lukper" then
             '*** Er ny lukepridoe i et interval der allerede er lukket?
                
                  strSQL = "SELECT lk_dato FROM lon_korsel WHERE lk_id <> 0 ORDER BY lk_dato DESC"
                 lastLkDato = "01-01-2002"
     
                 'Response.write strSQL & "<br>"

                 oRec.open strSQL, oConn, 3
                 if not oRec.EOF then

                 lastLkDato = oRec("lk_dato")

                 end if
                 oRec.close

                 lukDatoTjk = day(lukDato)&"-"&month(lukDato) &"-"& year(lukDato)
                 'Response.write "MM"& lukDatoTjk & ":: "& lastLkDato & "<br>"

             if cDate(lukDatoTjk) <= cDate(lastLkDato) then

                    %>
					<!--#include file="../inc/regular/header_hvd_inc.asp"-->
					     
			        <%
                    useleftdiv = "c"
					errortype = 164
					call showError(errortype)
		            Response.End

                  

             end if

             else

             lastLkDato = request("lastLkDato")
             
             end if


              if func = "lukper" then 'kun føste gang, da den ellers tjekker op med sin egen lønperide afslutning
             call lonKorsel_lukketPer(now, -2)
             else
             lonKorsel_lukketPerDt = lastLkDato
             end if

             '*** Periode er fra senest aflsuttet periode + 1 dag
              lonKorsel_lukketPerDt = dateAdd("d", 1, lonKorsel_lukketPerDt)

              %>

              <!--#include file="../inc/regular/header_hvd_inc.asp"-->
       
             <br /><br />
               <div style="position:relative; top:0px; left:40px; width:800px; padding:20px;">
             <img src="../ill/outzource_logo_200.gif" />
             </div>
             <div style="position:relative; top:20px; left:40px; width:800px; padding:20px;">

                 <%if cint(nulstil) = 1 then 'nulstil overfør ikke korrektioner / saldo

                  
                      '**** Lukker periode
                     strSQLluk = "INSERT INTO lon_korsel (lk_dato, lk_editor) VALUES ('"& lukDato &"', '"& session("user") &"')"
	                 oConn.execute(strSQLluk)
                  
          

                     %>
                   <a href="bal_real_norm_2007.asp?menu=stat"> Periode er afsluttet >></a><br /><br />

                 <%

             response.end

                  end if
             %>

             <h4>Overført til ny periode d. <%=formatdatetime(lukDato, 1)%>
             <br /><span style="font-size:9px; font-weight:lighter;">Print evt. til dokumentation, inden du fortsætter.</span></h4>

             <!--Henter medarbejder <%=replace(lmtSQL, ",", " - ")%>-->
             <table cellspacing=2 cellpadding=2 border=0 width=100%>
             <tr><td><b>Medarbejder</b></td><td>
             <%if func = "lukper_ok" then %>
             <b>Saldo Norm. / Realiseret</b>
             <%else %>
             &nbsp;
             <%end if %>
             </td><td>
             
              <%if func = "lukper_ok" then %>
             <b>Saldo Norm. / Komme/Gå</b>
             <%else %>
             &nbsp;
             <%end if %>
             </td><td>Periode</td><td>Indlæst?</td>
             
             
             </tr>
             <tr> 

    <%      


            response.flush

             if func = "lukper_ok" then
             strSQL = "SELECT mid FROM medarbejdere WHERE mansat = 1 AND ansatdato <= '"& lukDato &"' AND mid > "& lastMid &" ORDER BY mid LIMIT 7" '(mid <> 1 AND mid <> 12 AND mid <> 21 AND mid <> 5 AND mid <> 7 AND mid <> 8) 'AND mid = 7 AND (mid <> 1 AND mid <> 12 AND mid <> 21)
             else
             strSQL = "SELECT mid FROM medarbejdere WHERE mansat = 1 AND ansatdato <= '"& lukDato &"' ORDER BY mid"
             end if
             
             m = 0
             oRec7.open strSQL, oConn, 3
             while not oRec7.EOF

             m = m + 1 

             call meStamdata(oRec7("mid"))
             

             'if func = "lukper" then
             if cdate(meAnsatDato) >= cdate(licensstdato) then
             listartDato_GT = meAnsatDato
             else
             listartDato_GT = licensstdato
             end if

             lic_mansat_dt = listartDato_GT

             if cdate(lonKorsel_lukketPerDt) > cdate(listartDato_GT) then
             lnkStDatoKri = lonKorsel_lukketPerDt
             listartDato_GT = year(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& day(lonKorsel_lukketPerDt)
             else
             listartDato_GT = year(listartDato_GT) &"/"& month(listartDato_GT) &"/"& day(listartDato_GT)
             lnkStDatoKri = listartDato_GT
             end if
             
             'Response.write "cdate(lonKorsel_lukketPerDt): "&  cdate(lonKorsel_lukketPerDt) &" listartDato_GT: "& cDate(listartDato_GT) & " lukDato: "& lukDato & " lastLuk: "& lastLkDato &"<br>"
             'Response.end
             'else


             'listartDato_GT = request("listartDato_GT")
             'listartDato_GT = year(listartDato_GT) &"/"& month(listartDato_GT) &"/"& day(listartDato_GT)
             
                    if func = "lukper_ok" then
                    'Response.write "listartDato_GT: "& listartDato_GT & " lukDato: "& lukDato & "<br>"
                    'Response.end
                    call medarbafstem(oRec7("mid"), listartDato_GT, lukDato, 50, akttype_sel, 0)

                    end if
            
            'end if
         


    %>
    <tr>
    <td><%=meTxt %></td>
    <%
   

    %>
    
    <td>
    <%=realNormBal %>
    </td>
    <%
    
    if session("stempelur") <> 0 then %>
   <td>
     <%=normLontBal %>
    </td>
    <%end if %>

    <td style="white-space:nowrap;"><%=formatdatetime(listartDato_GT, 1) & " - "& formatdatetime(lukDato,1) %></td>
    <td align=right><%=gkTxt %></td>
    </tr>

            <%
                
                if func = "lukper_ok" then
                 call indlasKorrektioner(oRec7("mid"), normLontBal, realNormBal, lukDato_1)
                 end if
                 
                  

                  lastMid = oRec7("mid")

            response.flush
               oRec7.movenext
            wend
            oRec7.close

            %>
            <tr><td colspan=5>
            <%

                    if aktidKorrFundet = 0 AND func = "lukper_ok" then%>
					<!--#include file="../inc/regular/header_hvd_inc.asp"-->
					     
			        <%
                    useleftdiv = "c"
					errortype = 163
					call showError(errortype)
		            Response.End

                    end if
             

            %>
            </td></tr>


            <tr><td colspan=5 ><br />Overførte saldi kan findes på "Korrektion's" aktiviterne under interne job.<br /> Korrektioner er indsat på den <%=formatdatetime(lukDato_1, 1) %> (første dag i den nye periode)</td></tr>
            
            <%if func = "lukper" then %>
           
            <tr><td colspan=5 align=right><br /><a href="bal_real_norm_2007.asp?menu=stat&func=lukper_ok&FM_lk_dag=<%=lukdag%>&FM_lk_md=<%=lukmd%>&FM_lk_aar=<%=lukaar%>&lmt=<%=lmt %>&lastMid=0&lastLkDato=<%=lastLkDato %>">Det ser korrekt ud - Fortsæt >></a><br />
            <span style="color:#999999;">(af load hensyn overføres der 7 medarb. ad gangen)</span></td></tr>
            <%else %>


               <tr><td colspan=5 align=right>
               
               
              

               <%
               
               derfindesflere = 0
                strSQLtjk = "SELECT mid FROM medarbejdere WHERE mansat = 1 AND mansat <= '"& lukDato &"' AND mid > "& lastMid &" ORDER BY mid LIMIT 1"
                oRec7.open strSQLtjk, oConn, 3
                if not oRec7.EOF then
                derfindesflere = 1
                end if
                oRec7.close


                if derfindesflere = 1 then %>
               <br /><br /><a href="bal_real_norm_2007.asp?menu=stat&func=lukper_ok&FM_lk_dag=<%=lukdag%>&FM_lk_md=<%=lukmd%>&FM_lk_aar=<%=lukaar%>&lmt=<%=nextLmt %>&lastMid=<%=lastMid%>&lastLkDato=<%=lastLkDato %>">Fortsæt med de næste medarbejdere >></a> <br />
                    <span style="color:#999999;">(Opdater ikke denne side igen, da du så vil indlæse tallene dobbelt)</span><br /><br />
               <%
               afsFontSize=12
               else
               afsFontSize=12
               end if%>
               <br /><br />
               <br /><br />
               Der er ikke flere medarbejdere på listen.<br />
               <a href="bal_real_norm_2007.asp?menu=stat" style="color:red; font-size:<%=afsFontSize%>px;">Periode er afsluttet. Afslut indlæsninger og vend tilbage [X]</a><br /><br />

               
            
       </td></tr>
            <%end if %>
            </table>

            <%

      
      if func = "lukper_ok" AND lmt = 0 then
      '**** Lukker periode (interne -2 + Komme / Gå)
     strSQLluk = "INSERT INTO lon_korsel (lk_dato, lk_editor) VALUES ('"& lukDato &"', '"& session("user") &"')"
	 oConn.execute(strSQLluk)
     end if


	 Response.end

	 'Response.Redirect "bal_real_norm_2007.asp?menu=stat"
	 
	
    else
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	call showError(150)
	Response.end
	
	end if
	
	
	case "slet_lukper" 
     lk_id = request("lk_id")

	%>

        <!--#include file="../inc/regular/header_hvd_inc.asp"-->
             <br /><br />
               <div style="position:relative; top:0px; left:40px; width:800px; padding:20px;">
             <img src="../ill/outzource_logo_200.gif" />
             </div>
             <div style="position:relative; top:20px; left:40px; width:400px; padding:20px;">
             

             <h4>Slet afsluttet lønperiode?</h4>
      
             <table cellspacing=2 cellpadding=2 border=0 width=100%>
             <tr><td><a href="#" onclick="Javascript:history.back()"><< Nej, annuler</a></td>
        <td align=right><a href="bal_real_norm_2007.asp?func=slet_lukper_ok&lk_id=<%=lk_id%>">Ja - slet afsluttet lønperiode >> </a></td>
         
         </tr>
         </table>
         <br /><br />
         Hvis en lønperiode-afslutning slettes, sletter man samtidig alle korrektioner på denne dato.
            </div>

    <%
     
    case "slet_lukper_ok" 

     lk_id = request("lk_id")
	 
     strSQL = "SELECT lk_dato FROM lon_korsel WHERE lk_id = " & lk_id
     tDato = "2002-01-01"
     
     'Response.write strSQL & "<br>"

     oRec.open strSQL, oConn, 3
     if not oRec.EOF then

     tDato = oRec("lk_dato")

     end if
     oRec.close

     tDato = dateadd("d", 1, tDato)
     tDato = year(tDato) &"/"& month(tDato) &"/"& day(tDato)


      strSQLdeltim = "DELETE FROM timer WHERE (tfaktim = 113 OR tfaktim = 114) AND tdato = '"& tDato & "'" 'AND editor = 'Korrektions indlæsning'
	  
      'Response.write strSQLdeltim
      'Response.end
      oConn.execute(strSQLdeltim)

      
	  strSQLdel = "DELETE FROM lon_korsel WHERE lk_id = " & lk_id
	  oConn.execute(strSQLdel)


	  
	  Response.Redirect "bal_real_norm_2007.asp?menu=stat"



    
   
	
	case else
	
	
    
	
  



	'dim intMids
	'redim intMids(150)
	
	


	'if len(request("FM_usedatokri")) <> 0 OR media = "print" OR media = "export" then
	
	 '   if left(request("FM_medarb"), 1) = 0 then
	 '   thisMiduse = "0"
   ' 	intMids = split(thisMiduse, ",")
	'    else
	'    thisMiduse = request("FM_medarb")
	 '   intMids = split(request("FM_medarb"), ",")
	 '   end if
	
	'else
	    
	'    thisMiduse = session("mid") '& ","
	'    intMids = split(thisMiduse, ",")
	   
	'end if
	
    if request("dontdisplayresult") <> "" then
    dontdisplayresult = 1 
    else
    dontdisplayresult = 0
    end if
	

    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if


	'*** Rettigheder på den der er logget ind **'
    'medarbid = session("mid")
	if len(request("FM_medarb")) <> 0 AND (media <> "export" AND media <> "print") then 
        
        if len(trim(request("FM_mdoversigt"))) <> 0 then
        mdoversigt = 1
        mdoversigtCHK = "CHECKED"
        else
        mdoversigt = 0
        mdoversigtCHK = ""
        end if

        response.Cookies("tsa")("c_mdoversigt") = mdoversigt


        if len(trim(request("FM_mdoversigt_ultimo"))) <> 0 then
        mdoversigt_ultimo = 1
        mdoversigtUltimoCHK = "CHECKED"
        else
        mdoversigt_ultimo = 0
        mdoversigtUltimoCHK = ""
        end if

        response.Cookies("tsa")("c_mdoversigt_u") = mdoversigt_ultimo



    else


        if request.Cookies("tsa")("c_mdoversigt") = "1" then
        mdoversigt = 1
        mdoversigtCHK = "CHECKED"
        else
        mdoversigt = 0
        mdoversigtCHK = ""
        end if


        if request.Cookies("tsa")("c_mdoversigt_u") = "1" then
        mdoversigt_ultimo = 1
        mdoversigtUltimoCHK = "CHECKED"
        else
        mdoversigt_ultimo = 0
        mdoversigtUltimoCHK = ""
        end if


    end if


	if len(request("FM_medarb")) <> 0 then 'OR func = "export"
	
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
	
	
	'Response.Write thisMiduse
	'Response.flush
	
	
	if len(request("FM_akttype")) <> 0 AND len(request("FM_usedatokri")) <> 0 then 
	akttype_sel = request("FM_akttype")
	else
	    if request.Cookies("tsa")("fm_akttype_sel") <> "" then
	    akttype_sel = request.Cookies("tsa")("fm_akttype_sel")
	    'akttype_sel = "#-99#, #-1#, #-2#, #-3#, #-4#, #-5#, "
	    else
	    'call akttyper2009(6)
	    akttype_sel = "#-99#, #-1#, #-2#, #-3#, #-4#, #-5#, "
	    akttype_sel = akttype_sel & aktiveTyper
	    end if
	end if
	
	
	'Response.Write akttype_sel
	response.Cookies("tsa")("fm_akttype_sel") = akttype_sel
	
	if trim(request("FM_visnulfilter")) <> "" OR len(request("FM_usedatokri")) <> 0 then
	visnulfilter = request("FM_visnulfilter")
        if visnulfilter <> "" then
        visnulfilter = 1
        else
        visnulfilter = 0
        end if
	else
	    if request.cookies("tsa")("visnulfilter") <> "" then
	    visnulfilter = request.cookies("tsa")("visnulfilter")
	    else
	    visnulfilter = 0
	    end if
	end if
	
	response.Cookies("tsa")("visnulfilter") = visnulfilter


    if trim(request("FM_visikkeFerieogSygiPer")) <> "" OR len(request("FM_usedatokri")) <> 0 then
	visikkeFerieogSygiPer = request("FM_visikkeFerieogSygiPer")
        if visikkeFerieogSygiPer <> "" then
        visikkeFerieogSygiPer = 1
        else
        visikkeFerieogSygiPer = 0
        end if
	else
	    if request.cookies("tsa")("visikkeFerieogSygiPer") <> "" then
	    visikkeFerieogSygiPer = request.cookies("tsa")("visikkeFerieogSygiPer")
	    else
	    visikkeFerieogSygiPer = 0
	    end if
	end if
	
	response.Cookies("tsa")("visikkeFerieogSygiPer") = visikkeFerieogSygiPer


 



     korrOversktiftIsWrt = 0
	
	if media <> "export" then
	
	
	if media <> "print" then
	
	leftPos = 90
	topPos = 102
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	 <div id="loadbar" style="position:absolute; display:; visibility:visible; top:350px; left:380px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid: 5-20 sek.
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>
	
	
	
	<script src="inc/bal_real_norm_jav.js"></script>
	

 
	<%
        
        call menu_2014()
        
        else 
	
	leftPos = 20
	topPos = 20
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if%>
	
	<%end if
    
    showextended = 0
    call stempelur_kolonne(lto, showextended)%>
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">
	
	
	<%
	
    'if media = "print" then
	'oimg = "ikon_medarbaf_48l.png"
	oleft = 0
	otop = -30
	owdt = 600
	oskrift = "Medarbejder-afstemning & Løn (HR listen)"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	'end if
	
	
	  '*** Licens dato ****'
      call licensStartDato()
	 public lisStDato
	 lisStDato = startDatoDag&"/"&startDatoMd&"/"&startDatoAar 


       '** Finder seneste afsluttetede periode ***'
	    lastLonper = startDatoDag&"-"&startDatoMd&"-"&startDatoAar
	    
	    strSQLlukper = "SELECT lk_id, lk_dato, lk_editor FROM lon_korsel WHERE lk_id <> 0 ORDER BY lk_dato DESC LIMIT 1"
        oRec.open strSQLlukper, oConn, 3
        h = 0
        if not oRec.EOF then
        
        lastLonper = oRec("lk_dato")
        
        end if
        oRec.close


        '*** Ultimo Saldo ***************************** 
        if cDate(lastLonper) > cDate(lisStDato) then
        ultimoStDato = lastLonper 
        else
        ultimoStDato = lisStDato
        end if 

        ultimoSlDato = "1-"& month(now) &"-"& year(now)
        ultimoSlDato = dateAdd("d", -1, ultimoSlDato)



	if media <> "print" AND media <> "export" then
	
	call filterheader_2013(0,0,1004,oskrift)%>
	<table border=0 cellspacing=0 cellpadding=10 width=100%>
	<form action="bal_real_norm_2007.asp" method="post" name="periode" id="periode">
    <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
	<tr>
	

    <%call progrpmedarb %>


  
	
	
	<td valign=top><b>Periode:</b><br />
	<%
    response.Write "Licens startdato:<b> "& formatdatetime(lisStDato, 1) & "</b><br>"
	%>
	 
	 
	Seneste lønperiode afsl.: <b><%=formatdatetime(lastLonper, 1) %></b><br /><br />
    
  

	<!--#include file="inc/weekselector_s.asp"--> 
	<%
	startdato = strDag&"/"&strMrd&"/"&strAar
	slutdato = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
	%>

    <br /><br />
    <table width=80% cellpadding=0 cellspacing=0 border=0>
    <tr><td width=50%><input id="addDag_minus" type="button" value=" << - Dag " style="font-size:9px;" />
    </td>
    <td align=right><input id="addDag" type="button" value=" + Dag >> " style="font-size:9px;" />
    </td>
    </tr>
    

    <input id="addDagVal" name="addDagVal" value="0" type="hidden" />
    </table>

        <br /><br />
        <input type="checkbox" name="FM_mdoversigt" <%=mdoversigtCHK %> value="1" /> Vis rapport som månedsoversigt<br />
        
             <span style="color:#999999; font-size:9px;">Månedsopdelt fra d. 1 i valgte måned til sidste dag i den valgte slut måned. <b>Maks 12 måneder</b>.</span>
        
        <br /><br /><input type="checkbox" name="FM_mdoversigt_ultimo" id="FM_mdoversigt_ultimo" <%=mdoversigtUltimoCHK %> value="1" /> Vis saldo ultimo <%=monthname(month(dateAdd("m", -1, now))) &" "& year(dateAdd("m", -1, now)) %> <br />
         <span style="color:#999999; font-size:9px;">Henter saldo fra licensstart / seneste afl. lønperiode til ultimo seneste måned.  Husk at afslutte lønperioder for bedre loadtid</span>
    

	</td>
	</tr>
	<tr><td colspan=3>
        
	<%if visnulfilter = 1 then
	visnulfilterCHK = "CHECKED"
	else 
	visnulfilterCHK = ""
	end if%>
        <input id="FM_visnulfilter" name="FM_visnulfilter" value="1" type="checkbox" <%=visnulfilterCHK %> />Skjul "nul" linier (medarbejdere uden timer på valgte typer)


        <%if cint(visikkeFerieogSygiPer) = 1 then
	visikkeFerieogSygiPerCHK = "CHECKED"
	else 
	visikkeFerieogSygiPerCHK = ""
	end if%>
        <br /><input id="FM_visikkeFerieogSygiPer" name="FM_visikkeFerieogSygiPer" value="1" type="checkbox" <%=visikkeFerieogSygiPerCHK %> />Skjul Ferie, sygdom og overarb. i periode kolonner. (Vis kun Ferie i ferieår og Sygdom ÅTD)

	    </td></tr>
	<tr>
	<td colspan=3>
	
	<%response.flush %>
	
	<br /><b>Vis kolonner for følgende aktivitets typer:</b> (Vælg)<br />
			<table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr><td valign=top style="padding:3px; width:125px; border:1px #D6DFf5 solid;">
		
			<b>Rapporttyper:</b><br />
			
			<table cellspacing=2 cellpadding=1 border=0><tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_lon" class=vmenu>Løn</a></td></tr>
			<tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_ferie" class=vmenu>Ferie</a></td></tr>
			<tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_afstem" class=vmenu>Afstemning</a></td></tr>
			<tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_udspec" class=vmenu>Afstemning + Udspecificering</a></td></tr>
			<tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_syg" class=vmenu>Fravær & Sygdom</a></td></tr>
			<tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_all" class=vmenu>Alt</a></td></tr>
            <%if lto = "fk" OR lto = "intranet - local" then %>
            	<tr><td style="border-bottom:1px #999999 solid;"><a href="#" id="rap_fk_sdlon" class=vmenu>FK lønrap. SD løn</a></td></tr>
            <%end if %>

			</table>
			
			</td><td valign=top style="padding:3px; width:125px; border:1px #D6DFf5 solid;">
			<%
			call akttyper2009(5)
			%>
			</table>
			</td></tr>
			</table>
	
	
	
	</td>
	    
		
	
	
	</tr>
    <tr>    
	    
	    <td colspan=3 align=right valign=bottom style="padding-top:16px;">


            <input id="Submit1" type="submit" value=" Søg >> " />
	</td></tr>
    </form>
    <!--
    <form>
	  <tr><td colspan=3><b>Abonnér</b><br />
    <input type="checkbox" /> Ugentligt<br />
    <input type="checkbox" /> Månedligt<br />
    Navn på rapport: <input type="text" style="width:200px;" />
    <input type="submit" value="Gem" /></td></tr>
    </form>
    --></table>
    
    
    
    <!-- filter header sLut -->
	</td></tr></table>
	</div>
    
    
    <%
    
    else
    dontshowDD = 1
    %>
    <!--#include file="inc/weekselector_s.asp"--> 
	<%
	startdato = strDag&"/"&strMrd&"/"&strAar
	slutdato = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
	
    'Response.Write "Periode: "&  formatdatetime(startdato, 1) & " til "& formatdatetime(slutdato, 1) & "<br><br>"
    
    end if 'print
    
    
    
			

if media <> "print" AND media <> "export" then

'Response.flush

ptop = 0
pleft = 1034
pwdt = 200

            call eksportogprint(ptop,pleft,pwdt)

            if cint(dontdisplayresult) <> 1 then 
            %>


             <form action="bal_real_norm_2007.asp?media=export" method="post" target="_blank">
                 <input id="Hidden2" name="FM_medarb" value="<%=thisMiduse%>" type="hidden" />
                 <input id="Hidden1" name="FM_medarb_hidden" value="<%=thisMiduse%>" type="hidden" />
            <tr>
    
  
                <td align=center><input type=image src="../ill/export1.png" /></td>
                <td><input id="Submit4" type="submit" value=".csv fil eksport" style="font-size:9px; width:120px;" />
                 </td>
 
                </tr>
   
                <%select case lto
                case "intranet - local", "fk"
                if level = 1 then %>
                 <tr><td colspan=2 valign=top>
                <input type="checkbox" name="sd_lon_fil" value="1" /> Tilknyt SD lønfil fra CSV fil.
                </td></tr>
                <% end if
                case else
                end select %>
   
            </form>

             <form action="bal_real_norm_2007.asp?media=print" method="post" target="_blank">
                 <input id="Hidden3" name="FM_medarb" value="<%=thisMiduse%>" type="hidden" />
                 <input id="Hidden4" name="FM_medarb_hidden" value="<%=thisMiduse%>" type="hidden" />
            <tr>
    
  
                <td align=center><input type=image src="../ill/printer3.png" /></td>
                        <td><input id="Submit3" type="submit" value="Print version" style="font-size:9px; width:120px;" /></td>
 
                </tr>
            </form>

   
   
  
                <tr>
                <td align=center><a href="#" onclick="Javascript:window.open('bal_real_norm_2007.asp?sendemail=j&FM_medarb=<%=thisMiduse%>', '', 'width=100,height=100,resizable=no,scrollbars=no')" class=rmenu>
               &nbsp;<img src="../ill/ikon_sendemail_24.png" border=0 alt="" /></a>
            </td><td>
                <a href="#" onclick="Javascript:window.open('bal_real_norm_2007.asp?sendemail=j&FM_medarb=<%=thisMiduse%>', '', 'width=100,height=100,resizable=no,scrollbars=no')" class=rmenu>Email til valgte medarb.</a>
            </td>
               </tr>
               
                 <tr><td>&nbsp;</td><td>
                     
                     <%if level = 1 then
	                lnktxt = "Ferie, Afspad. & Sygdom's kalender"
	                else
	                lnktxt = "Ferie & Afspad. kalender"
	                end if
	                %>
	
	            <a href="feriekalender.asp?menu=job" class=rmenu><%=lnktxt %> >></a></td></tr>

                <%else %>
                <tr><td><span style="color:#999999;"><i>Afventer indhold..</span></i><br /><br /><br><br /><br>&nbsp;</td></tr>
                 <% end if'dontdisplayresult%>
   
	
   </table>
 
</div>




                    <%if level = 1 then%>

                    <%
                    select case lto
                    case "intranet - local", "fk"
                    tTop = 0
                    case else
                    tTop = 0
                    end select

                    tLeft = 1265
                    tWdth = 225
                    tHgt = 658 
                    tId = "lk_div"
                    tVzb = "visible"
                    tDsp = ""
                    tZindex = 1000

                    call tableDivAbs(tTop,tLeft,tWdth,tHgt,tId, tVzb, tDsp, tZindex)



                    %>


                    <FORM method="post" action="bal_real_norm_2007.asp?func=lukper">
                    <table cellspacing=0 cellpadding=0 border=0 width=100%>

                    <tr>
	                    <td style="padding:13px 10px 10px 10px;"><h4>Afslut Lønperiode</h4>
                        Husk at ajourf&oslash;re afsluttede l&oslash;nperioder. == bedre loadtid på afstemning. <br /><br />
                        Interne job (HR) og Komme/Gå bliver lukket for registrering og der bliver overført saldo (korrektion) til den nye periode.
                        </td>
	                    </tr>
                    <td style="padding:10px 10px 10px 10px;">	
	
                            <b>Afslut periode dato:</b><br />
                            <input id="FM_lk_dag" name="FM_lk_dag" type="text" value="<%=day(now) %>" style="width:20px; font-size:9px; font-family:arial;" /> - 
                            <input id="FM_lk_md" name="FM_lk_md" type="text" value="<%=month(now) %>" style="width:20px; font-size:9px; font-family:arial;" /> - 
                            <input id="FM_lk_aar" name="FM_lk_aar" type="text" value="<%=year(now) %>" style="width:30px; font-size:9px; font-family:arial;" />   dd - mm - åååå 

                            <br /><input type="checkbox" value="1" name="FM_nulstil" /> Nulstil ny periode <span style="color:#999999;">(overfør ikke flekssaldo fra gl. periode)</span><br /><br />

                            <input id="Submit2" type="submit" value="Afslut periode >>" />
                          
                           <br /><br />
                            <b>Historik</b> (seneste 12)<br />
        
                            <table cellspacing=0 cellpadding=1 border=0 width=100%>
                            <tr bgcolor="#D6Dff5">
                                <td class=lille><b>Dato</b></td>
                                <td class=lille><b>Afslut. af</b></td>
                                <td>&nbsp;</td>
                            </tr>
                            <%
                            strSQLlukper = "SELECT lk_id, lk_dato, lk_editor FROM lon_korsel WHERE lk_id <> 0 ORDER BY lk_dato DESC LIMIT 12"
                            oRec.open strSQLlukper, oConn, 3
                            h = 0
                            while not oRec.EOF 
        
                            select case right(h, 1)
                            case 0,2,4,6,8
                            bgcol_lk = "#EFf3FF"
                            case else
                            bgcol_lk = "#FFFFFF"
                            end select
        
                            %>
                            <tr bgcolor="<%=bgcol_lk %>"><td class=lille><%=oRec("lk_dato") %></td>
                                <td class=lille><i><%=left(oRec("lk_editor"), 10) %></i></td>
                                <td><a href="bal_real_norm_2007.asp?func=slet_lukper&lk_id=<%=oRec("lk_id") %>" class=red>[x]</a></td>

            
                            </tr>
                            <%
        
                            h = h + 1
                            oRec.movenext
                            wend
                            oRec.close
        
        
                             %></table>

        
	                    </td>
	                    </tr>

                    </table>
                    </form>
                    </div>


      

                    <%else%>


<%end if%>

<%end if%>


<% 
    if media = "print" then
        Response.Write("<script language=""JavaScript"">window.print();</script>")
    end if
    
    

    strEksportTxtMd = ""

    if cint(dontdisplayresult) <> 1 then 

    if media <> "print" then
	tTop = 10
    else
    tTop = -20
    end if
	
    tLeft = 0
	tWdth = 800
	globalWdt = 250 
    glbWdtCount = split(akttype_sel, ",")
        
        for t = 0 to UBOUND(glbWdtCount)
        globalWdt = globalWdt + 110
        next

	'startdatoSQL = strAar&"/"&strMrd&"/"&strDag
	'slutdatoSQL = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
        if media <> "print" then
        call tableDiv(tTop,tLeft,tWdth)
        else
        call tableDiv_j(tTop,tLeft,tWdth)
        end if


        'Response.write "mdoversigt: "& mdoversigt
	
    if cint(mdoversigt) = 1 OR cint(mdoversigt_ultimo) = 1 then
    endp = dateDiff("m", startdato, slutdato, 2,2)
        
        if cint(mdoversigt) = 1 then
        stPoint = 0
        else
        stPoint = endp + 1
        end if
        
        if mdoversigt_ultimo = 1 then
        endp = endp + 1
        end if

        if endp > 12 then

        if cint(mdoversigt_ultimo) = 1 then
        endp = 13

            if cint(mdoversigt) = 1 then
            stPoint = stPoint
            else
            stPoint = endp
            end if

            
        else
        endp = 12
        end if
        
        end if

    else
    stPoint = 0
    endp = 0
    end if


     
      headerwrtExp = 0

     
    for p = stPoint to endp

         

         headerwrt = 0
      


        if endp <> 0 then
        

        if p <> endp then 
        monthDt = dateadd("m",p,startdato)
        startdatoDt = "1-"& month(monthDt) &"-"& year(monthDt)
        
        slutdatoDt = dateAdd("m", 1, startdatoDt)
        slutdatoDt = dateAdd("d", -1, slutdatoDt)
        'slutdatoDt = year(slutdatoDt) &"-"& month(slutdatoDt)&"-"& day(slutdatoDt)
            if media <> "print" then 
            Response.write "<br><br><h4>" & monthname(datePart("m", monthDt, 2,2)) & " "& datePart("yyyy", monthDt, 2,2) &"</h4>"
            else
            Response.write "<br><br><b>" & monthname(datePart("m", monthDt, 2,2)) & " "& datePart("yyyy", monthDt, 2,2) &"</b><br>"
            end if
        

             if media = "export" then
           
            strEksportTxtMd = "xx99123sy#z xx99123sy#z" & monthname(datePart("m", monthDt, 2,2)) & " "& datePart("yyyy", monthDt, 2,2) 
            
            strEksportTxt = strEksportTxt & strEksportTxtMd
            end if
           
       
        else
                
               

                if cint(mdoversigt_ultimo) = 1 then
                'monthDt = dateadd("m",p,startdato)
                startdatoDt = ultimoStDato
                slutdatoDt = ultimoSlDato 'dateAdd("m", 0, slutdato)
                'slutdatoDt = dateAdd("d", -1, slutdatoDt)
                  
                if media <> "print" then 
                Response.write "<br><br><h4><span style=""font-size:9px;"">Saldo ultimo:</span><br>"& monthname(datePart("m", slutdatoDt, 2,2)) & " "& datePart("yyyy", slutdatoDt, 2,2) &"</h4>"
                else
                
                if cint(mdoversigt) = 1 then    
                Response.write "<br><br>"
                end if

                Response.write "Saldo ultimo: <b>"& monthname(datePart("m", slutdatoDt, 2,2)) & " "& datePart("yyyy", slutdatoDt, 2,2) &"</b><br>"
                
                end if
                'Response.write "ultimoStDato:" & ultimoStDato & " ultimoSlDato: "& ultimoSlDato
                          

                else

                monthDt = dateadd("m",p,startdato)
                startdatoDt = "1-"& month(monthDt) &"-"& year(monthDt)
        
                slutdatoDt = dateAdd("m", 1, startdatoDt)
                slutdatoDt = dateAdd("d", -1, slutdatoDt)
                'slutdatoDt = year(slutdatoDt) &"-"& month(slutdatoDt)&"-"& day(slutdatoDt)
                
                        if media <> "print" then 
                        Response.write "<br><br><h4>" & monthname(datePart("m", monthDt, 2,2)) & " "& datePart("yyyy", monthDt, 2,2) &"</h4>"
                        else
                         Response.write "<br><br><b>" & monthname(datePart("m", monthDt, 2,2)) & " "& datePart("yyyy", monthDt, 2,2) &"</b><br>"
                        end if

                        
                         if media = "export" then
                         strEksportTxtMd = "xx99123sy#z xx99123sy#z" & monthname(datePart("m", monthDt, 2,2)) & " "& datePart("yyyy", monthDt, 2,2) 
                         strEksportTxt = strEksportTxt & strEksportTxtMd
                         end if

                end if


        end if


        else

        startdatoDt = startdato
        slutdatoDt = slutdato

        if media <> "print" then 
        Response.write  "<br><br><h4><span style=""font-size:10px;"">Periode afgrænsning:</span><br> "& formatdatetime(startdato, 1) & " - "&  formatdatetime(slutdato, 1) & "</h4>"
        else
        Response.write  "<br><b><span style=""font-size:10px;"">Periode afgrænsning:</span><br> "& formatdatetime(startdato, 1) & " - "&  formatdatetime(slutdato, 1) & "</b>"
        end if

            if media = "export" AND request("sd_lon_fil") <> "1" then
            strEksportTxtMd = "xx99123sy#z xx99123sy#zPeriode afgrænsning: "& formatdatetime(startdato, 1) & " - "&  formatdatetime(slutdato, 1)
            strEksportTxt = strEksportTxt & strEksportTxtMd
            end if

        end if
	
	
	
	'Response.Write "<br><br>mdoversigt:"& mdoversigt &" startdato:"& startdato &"/" & startdatoDt & " # "& slutdatoDt & "P. "& p & " endP:"& endp & " mdoversigt_ultimo "& mdoversigt_ultimo
	
	'** Timer Realiseret ***'
	'if len(request("FM_usedatokri")) <> 0 then
	for m = 0 TO UBOUND(intMids) '0=##
	    if intMids(m) <> 0 then
	    call medarbafstem(intMids(m), startdatoDt, slutdatoDt, 1, akttype_sel, m)
	    end if
	    if media <> "export" then
	    Response.flush
	    end if
	next
        m = 0
	'else
	'Response.Write "<tr><td><b>Medarbjeder afstemning</b><br>Vælg de ønskede medarbejdere i listen ovenfor...</td></tr>"
	'end if
    %>

	</table>
 
    <% if media <> "export" then
        response.Flush
        end if
         next %>

 <!-- table div --> 


</div>
<%end if %>


<form><input type="hidden" id="globalWdt" value="<%=globalWdt %>" /></form>
    
	<%
    if media <> "print" ANd media <> "export" then 
    %>
    <br />
	<table><tr><td class=lille style="color:#999999;">
	
        <%
			tloadB = now
	        Response.flush
        	loadtid = datediff("s", tloadA, tloadB, 2,2)
            Response.Write "Sidens loadtid: " & formatnumber(loadtid, 2) & " sek."
		
			 %>
	</td></tr></table>
	

	
	<br /><br />
    <%end if %>


	
	<%if media = "export" then 

   
	ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	call TimeOutVersion()
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\bal_real_norm_2007.asp" then
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
				'objF.WriteLine(strOskrifter & chr(013))
				objF.WriteLine(ekspTxt)
				

                '** FK opret eksport fil til SD løn ***'
                if len(trim(request("sd_lon_fil"))) <> 0 then
                %>
                <!--#include file="eksport_timeout_sdlon.asp"-->
                <%
                Response.redirect OutputFileName
                else
                Response.redirect "../inc/log/data/"& file &""	
                end if 
				
				
				
	
	

end if %>

     
      <%if sendemail = "j" then %>
	  <% 
	    Response.Write("<script language=""JavaScript"">window.alert(""Der er blevet afsendt en email til de valgte medarbejdere"");</script>")
        Response.Write("<script language=""JavaScript"">window.close();</script>")
        %>
	  <%end if %>
	
	
	  <%if media <> "print" ANd media <> "export" then  
        itop = -16
        ileft = 825
        iwdt = 120
        ihgt = 0
        ibtop = 10 
        ibleft = 150
        ibwdt = 600
        ibhgt = 1150
        iId = "pagehelp"
        ibId = "pagehelp_bread"
        call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)%>		
       
   
			 <br /><br />
    
			<%
			call akttyper2009(3)
			%>
			</table>
			
	
	    
	    <br /><br />
	    <b>Ferie</b><br />
	    Fra den 1. januar 2001 optjenes 2,08 dages ferie for hver måneds ansættelse i optjeningsåret, som er lig kalenderåret. 
        Dette gælder også medarbejdere på del-tid.<br />
        2,08 * 12 måneder = 25 dage ell. 5 arbejdsuger.
	    <br /><br />
	    <b>Dage</b><br />
	    Ferie of ferie fridage (afspad.) timer angives som timer og omregnes til dage udfra normeret timer pr. uge 
	    (total timer pr. uge / med 5 arb. dage, se medarbejdertyper ell. medarb. afstemning).<br /><br />
	    Hvis man er ansat 37 timer, 
	    angiver man altså 7,4 timer for en hel dags ferie, og 37 timer for en uge. 
	    
	    <br /><br />
	    <b>Fravær i periode</b><br />
	     (Norm. timer / ikke optjent fravær dvs. Syg + Barnsyg)<br />

	    
	    <br />
	    
                &nbsp;
	   
	    <!-- side info slut -->
        </td></tr></table>
			</div>


			<br /><br />&nbsp;
			
		<%end if %>	
	    
	
			
			<%if media <> "print" ANd media <> "export" then  %>
			<br /><br /><br />
             &nbsp;
            <%end if %>
			</div>
	
	<% 
	end select
	
	
	
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
