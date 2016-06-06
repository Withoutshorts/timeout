<%response.buffer = true %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%


    

 '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_sogjobogkunde"


                '*** SØG kunde & Job            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                if filterVal <> 0 then
            
                 lastKid = 0
                
                  strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
                         


                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) AND "_
                &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '"& jobkundesog &"%' OR "_
                &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    

                 'response.write "strSQL " &strSQL
                 'response.end

                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                if lastKid <> oRec("kid") then
                strJobogKunderTxt = strJobogKunderTxt &"<br><br><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b><br>"
                end if 
                 
                strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                strJobogKunderTxt = strJobogKunderTxt & "<input type=""checkbox"" class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &"> "& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"<br>" 
                
                lastKid = oRec("kid") 
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt


                    response.write strJobogKunderTxt

                end if    


        
          case "FN_sogakt"

               
                '*** Søg Aktiviteter 
                

                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                aktsog = request("jq_newfilterval")
                else
                filterVal = 0
                aktsog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")
                aktid = request("jq_aktid")
    
                if len(trim(request("jq_jobid"))) <> 0 then        
                jobid = request("jq_jobid")
                else
                jobid = 0
                end if

                'positiv aktivering
                pa = request("jq_pa") 



                if filterVal <> 0 then
            
                 
    
                strAktTxt = strAktTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_aktsog"">[X]</span>"    
                         

               if pa = 1 then
               strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               &" WHERE tu.medarb = "& usemrn &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"   


               else


                        '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& usemrn & " GROUP BY projektgruppeId"

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 


           




               strSQL= "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"      
    

                 'response.write "strSQL " &strSQL
                 'response.end
            
                end if


                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                 showAkt = 0
                if instr(medarbPGrp, "#"& oRec("projektgruppe1") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe2") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe3") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe4") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe5") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe6") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe7") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe8") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe9") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe10") &"#") <> 0 then
                showAkt = 1
                end if 


                
                
                if showAkt = 1 then 
                 
                strAktTxt = strAktTxt & "<input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">"
                strAktTxt = strAktTxt & "<input type=""checkbox"" class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &"> "& oRec("aktnavn") &"<br>" 
                
                end if
                
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if    




        end select
        response.end
        end if




tloadA = now



    %>
    
	<%
	
    'select case func 
	
	

    'case else
        %>

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <SCRIPT language=javascript src="inc/ugeseddel_2011_jav.js"></script>

        <%
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


        

    call akttyper2009(2)
	
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
	
	%>

	

     <%call browsertype()
    
        
        if browstype_client <> "ip" then
            
            
	  
         
         
                 if cint(nomenu) <> 1 then
         
                   leftPos = 90
	            topPos = 102
         
         
                call menu_2014() %>
	
	
	              <% 
                   else   
              
                        leftPos = 20
	                    topPos = 20

                   end if


          else

          
	    leftPos = 20
	    topPos = 99

          end if

    
    else 
	
	leftPos = 20
	topPos = 20
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if 'print%>
	
	<%end if 'eksport%>
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">
        <%
        

         

    tTop = 2
	tLeft = 0
	tWdth = 900
	
                 

	
	call tableDiv(tTop,tLeft,tWdth)    
            
            
     if media <> "print" AND len(trim(strSQLmids)) > 0  then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16 %>
	<form id="filterkri" method="post" action="ugeseddel_2011.asp">
        <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
        <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
        <table>
       <tr bgcolor="#ffffff">
	<td valign=top> <b><%=tsa_txt_077 %>:</b> <br />
    <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)
   
	<br />
				<%
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat <> 2 "& strSQLmids &" GROUP BY mid ORDER BY Mnavn"
					
					'Response.Write strsQL
					'Response.flush

					%>
					<select name="usemrn" id="usemrn" style="width:250px;"><!-- onchange="submit(); -->
					<%
					
					oRec.open strSQL, oConn, 3
					while not oRec.EOF 
					
					if cint(oRec("Mid")) = cint(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if%>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn")%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>
        </td>
           </tr></table>
        </form>
        <br /><br />
	
	<%
   
       end if 'media
	
	   

	

    

    
   ugeseddelvisning = 1

   perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
   perIntervalLoop = perInterval

   'response.write "perIntervalLoop " & perIntervalLoop
   for l = 0 to perIntervalLoop 
        
        if l = 0 then
        varTjDatoUS_use = varTjDatoUS_man
        showheader = 1
        showtotal = 0
        else
        varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
        showheader = 0
        showtotal = 0
         end if 

       if l = perIntervalLoop then
       showtotal = 1
       end if

         varTjDatoUS_son = varTjDatoUS_man

        call ugeseddel(usemrn, varTjDatoUS_use, varTjDatoUS_use, ugeseddelvisning, showtotal, showheader)  
   
   next
	
	
    if browstype_client <> "ip" then

        if media <> "print" then

        %>
        <div style="position:absolute; top:0px; left:945px; width:200px; background-color:#FFFFFF; padding:0px 20px 20px 20px;">
        <h4>Ugeseddel Resumé<br /><span style="font-size:9px; font-weight:lighter;">Normtimer, realiseret og timer uden match</span></h4>

        <% call normRealGrafWeekPage(usemrn, varTjDatoUS_man)  %>
            






   <!-- hvis level = 1 OR teamleder -->
   <%if level <=2 OR level = 6 then %>

       <br /><br />
       <!--
       <div style="border:1px #CCCCCC solid; padding:2px; width:180px;">
       -->

              


      
              
               
               <%if cint(showAfsuge) = 0 then
               
               
                if cint(ugegodkendt) <> 1 then 
                
                              if cint(ugegodkendt) = 2 then

                             call meStamdata(ugegodkendtaf)%>
                        <div style="color:#FFFFFF; font-size:11px; background-color:#FF6666; padding:5px;"><b>Ugeseddel er afvist!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#FFFFFF;"><i><%=ugegodkendtdt %> af <%=meNavn %></i></span>
                        <%if len(trim(ugegodkendtTxt)) <> 0 then %>
                        <br />
                        <span style="font-size:9px; line-height:12px; color:#000000;"><i><%=left(ugegodkendtTxt, 200) %></i></span>
                        <%end if %>
                        </div>

                        <%end if %>

                           <form action="<%=fmLink%>&func=godkendugeseddel" method="post">
                               <table width=90% cellpadding=0 cellspacing=0 border=0>
                            <tr><td class=lille><br />

                           <span style="font-size:11px;"><b>Godkend ugeseddel</b></span><br />
                           Når en ugeseddel godkendes, godkendes alle ugens registreringer automatisk.
                
                           <input id="Submit2" type="submit" value="Godkend ugeseddel >>" style="font-size:9px; width:120px;" />
             
                           </td></tr></table>
                           </form>

                <%else 
                        
                        call meStamdata(ugegodkendtaf)%>
                        <div style="color:green; font-size:11px; background-color:#DCF5BD; padding:5px;"><b>Ugeseddel er godkendt!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#999999;"><i><%=ugegodkendtdt %> af <%=meNavn %></i></span></div>
                
                <%end if %>

               <%else %>
                
                 <form action="<%=fmLink%>&func=adviserugeafslutning" method="post">
                <table width=90% cellpadding=0 cellspacing=0 border=0>
                <tr><td class=lille>

                  <span style="font-size:11px;"><b>Godkend ugeseddel</b></span><br />
                 Ugeseddel kan IKKE godkendes før den er afsluttet af medarbejder.<br />
                
            

                <%if len(trim(request("showadviseringmsg"))) <> 0 then %>
                <br />
                  <div style="color:#000000; font-size:11px; background-color:#DCF5BD; padding:5px;"><b>Besked afsendt!</b></div>
                <%else %>
                  <br />
                  <b>Send email</b> med besked om at uge mangler af blive aflsuttet<br />
                <input id="Submit1" type="submit" value="Send besked >>" style="font-size:9px; width:120px;" />
                <%end if %>
           
             
               </td></tr></table>
                </form>
               <%end if 
               
               
               if cint(ugegodkendt) <> 2 AND cint(showAfsuge) = 0 then%>

               <form action="<%=fmLink%>&func=afvisugeseddel" method="post">
               <table width=90% cellpadding=0 cellspacing=0 border=0>
               <tr><td class=lille>
               <br />
               <span style="font-size:11px;"><b>Afvis ugeseddel</b></span><br />
               Begrundelse:<br />
               <textarea name="FM_afvis_grund" style="width:200px; height:40px;" /></textarea>
                <input id="Submit3" type="submit" value="Afvis ugeseddel >>" style="font-size:9px; width:120px;" /><br />
                 Afsender email til medarbejer om at ugeseddel er afvist, og åbner evt. allerede godkendt ugeseddel op igen.
                  </td></tr></table>
                 
               </form>

               <%end if %>


   
             

   
     <!--</div>-->
   
   <%end if 'level %> 
    
    

        </div><!-- ugesededel resume -->

        <% end if ''IP


        

    if media <> "print" then '<> print

ptop = 0
pleft = 1200
pwdt = 200

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&varTjDatoUS_son=<%=varTjDatoUS_son %>&media=print" method="post" target="_blank">
<tr> 
    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value=" Print venlig >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>


</table>

      

      

</div>
<%else

Response.Write("<script language=""JavaScript"">window.print();</script>")

end if 
    
%>

  

  
    <%end if 'print%>

  </div><!--tablediv-->
    <br /><br />&nbsp;


    </div><!--sidediv-->
    <br /><br />&nbsp;
    <%
	
	
	'end select
    'end if
	
	
	%>
<!--#include file="../inc/regular/footer_inc.asp"-->
