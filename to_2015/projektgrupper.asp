<%response.buffer = true 
 Session.LCID = 1030
  %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->




<%'** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
    response.end
    end if %>

<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->





<%
    func = request("func")

	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if


        if media <> "print" then    
        call menu_2014
    end if

    %>

        <div id="wrapper">
        <div class="content">

    <%
    select case func
    case "slet"
        '***Her spørges om det er ok at slette en medarbejder***

        oskrift = "Projektgrupper" 
        slttxta = "Du er ved at <b>SLETTE</b> en projektgruppe. Er dette korrekt?"
        slttxtb = "" 
        slturl = "projektgrupper.asp?menu=tok&func=sletok&id="&id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)


    case "sletok" 
    '*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM projektgrupper WHERE id = "& id &"")
	Response.redirect "projektgrupper.asp?menu=job"
                

   
    case "opdprorel"
	
	z = 0
	x = Request("FM_antal_x")
	y = Request("FM_antal_y")
	
	'Response.write "x " & x & " y " & y & "<br>"
	strProjektgruppeId = request("FM_projektgruppeId")
	
	'** Finder jobid på job hvor denne projktgruppe har adgang til ***'
    guidjobids = ""
    guideasyids = ""
	call projktgrpPaaJobids(strProjektgruppeId)
	
	call positiv_aktivering_akt_fn() 'wwf
                         
	
		            
	'oConn.execute("DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &"")
	
	
	For z = 0 to y - 1
	'Response.write z & "<br>"
	'Response.write "strMid: " & request("FM_Mid_"&z&"")& " , Optages: "& request("FM_medlem_"&z&"") & "<br>"
	'Response.flush
		
		strMid = request("FM_Mid_"&z&"")
		    
		    
		   if len(trim(strMid)) <> 0 then 
		    
		    fandtes = 0
		    strSQLfindes = "SELECT projektgruppeId, MedarbejderId"_
		    &" FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId & " AND MedarbejderId = "&strMid
		    
		    oRec4.open strSQLfindes, oConn, 3
		    if not oRec4.EOF then
		    
		    fandtes = 1
		    
		    end if
		    oRec4.close  
		    
		    'Response.Write "fandtes:" & fandtes
		    
		    '** Team leder **'
		    if request("FM_teamleder_"&z&"") = 1 then
    		teamleder = 1
    		else
    		teamleder = 0
    		end if
		    
            '** Notificer **'
		    if request("FM_notificer_"&z&"") = 1 then
    		notificer = 1
    		else
    		notificer = 0
    		end if


		    '** fandtes relation i forvejen ??***'
		    if cint(fandtes) = 1 then
		    
		        '*** Hvis ja, skal den slettes? ***'
		        if request("FM_medlem_"&z&"") <> "1" then
    		      oConn.execute("DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &" AND MedarbejderId = "& strMid &"")
                  'Response.Write " -- DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &"" & "<br

                  '*** Sletter fra timereg usejob og tilføjer igen **'
                  del = 1 '1 slet fra timereg_usejob / 0: tilføj kun
                  call tilfojtilTU(strMid, del)
                
                else
                 
                 oConn.execute("UPDATE progrupperelationer SET teamleder = "& teamleder &", notificer = "& notificer &" WHERE projektgruppeId = "& strProjektgruppeId &" AND MedarbejderId = "& strMid &"")  

                 end if
		    
		    else
		    
		         '*** skal den oprettes? ***
		         if request("FM_medlem_"&z&"") = 1 then
    		        
    		        
    		        
		            oConn.execute("INSERT INTO progrupperelationer (projektgruppeId, MedarbejderId, teamleder, notificer) VALUES ("& strProjektgruppeId &", "& strMid &", "& teamleder &", "& notificer &")")
		            
		            
		            '*** sætter aktiv i guiden aktiv job
                    'forvalgt = 0 'off / 1 = on
                    'Response.Write guideasyids
                    'Response.end
		            'call setGuidenUsejob(strMid, guidjobids, 1, guideasyids, forvalgt,2)
                    del = 0 '1 slet fra timereg_usejob / 0: tilføj kun
                    
                    if cint(positiv_aktivering_akt_val) <> 1 then
                    call tilfojtilTU(strMid, del)
		            end if
		            
		            
		         end if
		    
		    
		    
		    end if
		
		end if 'strmid <> 0
	
	Next
	
	               
	'Response.end
	
	Response.redirect "projektgrupper.asp?menu=job&lastid="&strProjektgruppeId

    %>
    
    <% case "med" 
        
     orgvir = 0   
     %>
    <!--Projektgruppe medlemmer-->
    <script src="js/projektgrupper_medl.js" type="text/javascript"></script>

    <div class="container">
        <div class="portlet">
        <div class="portlet-title"><h3><u>Projektgrupper (medlemmer)</u></h3></div>

        <form action="projektgrupper.asp?menu=job&func=opdprorel" method="post">
	    <input type="hidden" name="FM_projektgruppeId" value="<%=id%>">
        

        <div class="row">
        <div class="col-lg-10">&nbsp</div>
        <div class="col-lg-2"><button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button></div>
        </div>
            <%
	        strSQL = "SELECT id, navn, orgvir FROM projektgrupper WHERE id=" & id
	        oRec.open strSQL, oConn, 3
	        if not oRec.EOF then
	        gruppeNavn = oRec("navn")
            orgvir = oRec("orgvir")
	        end if
	        oRec.close

            
            medIAndreOrgProjGrpStr = "#0#"
            if cint(orgvir) = 1 then

            '*** Tjekker Om Medarb er med i andre projgrp
            strSQLmedIAndreOrgProjGrp = "SELECT p.id, m.mid FROM projektgrupper AS p "_
            &" LEFT JOIN progrupperelationer AS pr ON (pr.ProjektgruppeId = p.id) "_
            &" LEFT JOIN medarbejdere AS m ON (m.mid = pr.medarbejderid) "_
            &" WHERE orgvir = 1 AND p.id <> "& id & " AND mid IS NOT NULL "
              
                'response.write strSQLmedIAndreOrgProjGrp

                oRec.open strSQLmedIAndreOrgProjGrp, oConn, 3
                while not oRec.EOF  
             

                    medIAndreOrgProjGrpStr = medIAndreOrgProjGrpStr & ",#"& oRec("mid") &"#"
	                

                oRec.movenext
                wend 
                oRec.close


            end if

	        %>        

        <div class="portlet-body">
            <section>
            <p>Medarbejdere i projektgruppen <b><%=gruppeNavn%></b></p>
           
            <!--Liste med medlemmer--> 
            <table class="table dataTable table-striped table-bordered table-hover ui-datatable">                  
               <thead>
                   <tr>
                       <th style="width: 45%">Navn</th>
                       <th style="width: 20%">Status</th>
                       <th style="width: 5%">Tilføj/Fjern ?</th>
                       <th style="width: 5%">Teamleder ?</th>
                       <th style="width: 5%">Notificer</th>
                   </tr>
            </thead>
                    <% 
                    Dim MedarbId()

                    strSQL = "SELECT Mnavn, Mid, MedarbejderId, teamleder, init, mnr, mansat, notificer FROM progrupperelationer, medarbejdere"_
                    &" WHERE Mid = MedarbejderId AND ProjektgruppeId = "&id&" ORDER BY Mnavn"
	                oRec.open strSQL, oConn, 3
	                x = 0
	                Redim MedarbId(x)
	                while not oRec.EOF 
	
	               
	
	                Redim preserve MedarbId(x)
	                MedarbId(x) = oRec("Mid")
                  %>  
               
               <tbody>
                   <tr>
                       <td><a href="medarb.asp?func=red&id=<%=oRec("Mid") %>"><%=oRec("Mnavn")%> 
                           <% 
                            if len(trim(oRec("init"))) <> 0 then
                            %>
                            &nbsp;[<%=oRec("init") %>]
                            <%end if    
                            %>
                           </a></td>

                       <td>
                             <%if oRec("mansat") = "2" then %>
                            - Deaktiveret
                             <%end if %>

                             <%if oRec("mansat") = "3" then %>
                            - Passiv
                             <%end if %>

                       </td>

                       <td><input type="checkbox" name="FM_medlem_<%=x%>" value="1" CHECKED>
		               <input type="hidden" name="FM_Mid_<%=x%>" value="<%=oRec("Mid")%>"></td>
                       <%if cint(oRec("teamleder")) =  1 then
		                tmCHK = "CHECKED"
		                else
		                tmCHK = ""
		                end if %>
		
                        <%if cint(oRec("notificer")) =  1 then
		                ntCHK = "CHECKED"
		                else
		                ntCHK = ""
		                end if 
                       %>
                       <%if level = 1 then %>
		                <td><input type="checkbox" name="FM_teamleder_<%=x%>" value="1" <%=tmCHK %>></td>
                        <td><input type="checkbox" name="FM_notificer_<%=x%>" value="1" <%=ntCHK %>></td>
		                <%else%>
                       <td><input type="checkbox" name="" value="1" DISABLED <%=tmCHK %>>
                        <input id="Hidden1" name="FM_teamleder_<%=x%>" value="<%=oRec("teamleder") %>" type="hidden" /></td>
                        <td><input type="checkbox" name="" value="1" DISABLED <%=ntCHK %>>
                        <input id="Hidden1" name="FM_notificer_<%=x%>" value="<%=oRec("notificer") %>" type="hidden" /></td>
		                <%end if %>
                   </tr>
                   <%
                    x = x + 1
                   oRec.movenext
                    wend
                   oRec.close
                   %>
               </tbody>     
           </table>

            <!--Ikke medlemmer liste-->
            <div class="pad-t10"><p>Medarbejdere udenfor projektgruppen (aktive og pasive)</p></div>
            <table id="tb_progrp_ikkemed" class="table dataTable table-striped table-bordered table-hover ui-datatable">                  
               <thead>
                   <tr>
                       <th style="width: 45%">Navn</th>
                       <th style="width: 20%">Status</th>
                       <th style="width: 5%;"><input type="checkbox" value="0" id="FM_chk_all_add" /><br />Tilføj/Fjern ?</th>
                       <th style="width: 5%">Teamleder ?</th>
                       <th style="width: 5%">Notificer</th>
                   </tr>
                </thead>
                   <%

                    '*** Henter ikke medlemmer
	                Dim T
	                T = 0
	                For T = 0 to x - 1
	                strExclude = strExclude & "Mid <> "&MedarbId(T)&" AND "
	                Next
	
	                antChar = len(strExclude)
	                if antChar > 0 then
	                LantChar = left(strExclude, (antChar -5)) 
	                strExcludeFinal = "WHERE " & LantChar
	
	                strSQL = "SELECT Mnavn, Mid, init, mnr, mansat FROM medarbejdere "& strExcludeFinal &" AND (mansat <> '2') ORDER BY Mnavn"
	                else
	                strSQL = "SELECT Mnavn, Mid, init, mnr, mansat FROM medarbejdere WHERE mansat <> '2' ORDER BY Mnavn"
	                end if

                   
	                oRec.open strSQL, oConn, 3

                    y = x
	                r = oRec.recordcount

	                while not oRec.EOF
                    %>
             
               <tbody>
                   <tr>
                       <td><a href="medarb.asp?func=red&id=<%=oRec("Mid") %>"><%=oRec("Mnavn")%>
                           <% 
                            if len(trim(oRec("init"))) <> 0 then
                            %>
                            &nbsp;[<%=oRec("init") %>]
                            <%end if    
                            %>
                           </a></td>

                        <td>
                             <%if oRec("mansat") = "2" then %>
                            - Deaktiveret
                             <%end if %>

                             <%if oRec("mansat") = "3" then %>
                            - Passiv
                             <%end if %>

                       </td>

                    

                       
                           
                      <%
                       '** Må ikke være med i mere end 1 org. projekgruppe ***'   
                       if cint(orgvir) = 0 OR instr(medIAndreOrgProjGrpStr, ",#"& oRec("mid") &"#") = 0 then %>
                               
                           <td><input type="checkbox" name="FM_medlem_<%=y%>" class="chk_medlem_clas" value="1">
                          
		                   <input type="hidden" name="FM_Mid_<%=y%>" value="<%=oRec("Mid")%>"></td>
                           <%if level = 1 then %>
		                    <td><input type="checkbox" name="FM_teamleder_<%=y%>" value="1"></td>
                            <td><input type="checkbox" name="FM_notificer_<%=y%>" value="1"></td>
		                    <%else %>
		                    <td><input type="checkbox" name="" value="0" DISABLED>
                            <input id="Text1" name="FM_teamleder_<%=y%>" value="0" type="hidden" /></td>

                            <td><input type="checkbox" name="" value="0" DISABLED>
                            <input id="Text1" name="FM_notificer_<%=y%>" value="0" type="hidden" /></td>
		                    <%end if %>


                       <%else %>
                       <td colspan="3" style="font-size:9px;">Er allerede i Org. projektgruppe</td>

                       <%end if %>

                   </tr>
                   <%
                   y = y + 1
                   oRec.movenext
                   wend
                   oRec.close
                   %>
               </tbody>     
             </table>
            </section>

            <div style="margin-top:15px; margin-bottom:15px;">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                     
         
                    <div class="clearfix"></div>
                 </div>

        </div>
            <input type="hidden" name="FM_antal_y" value="<%=y%>">
       </form>
    </div>
    

</div>

<%
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
	
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if len(trim(request("FM_opengp"))) <> 0 then
		intopengp = 1
		else
		intopengp = 0
		end if

        if len(trim(request("FM_orgvir"))) <> 0 then
		orgvir = request("FM_orgvir")
		else
		orgvir = 0
		end if


        
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO projektgrupper (navn, editor, dato, opengp, orgvir) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intopengp &", "& orgvir &")")
		
			strSQL = "SELECT id FROM projektgrupper WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			
			lastid = oRec("id")
			
			end if
			
			oRec.close
		
		else
		
		oConn.execute("UPDATE projektgrupper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', opengp = "& intopengp &", orgvir = "& orgvir &" WHERE id = "&id&"")
		lastid = id
		
		end if


        if func = "dbopr" AND request("FM_tilfoj_m") = "1" then
        '*** tilføjer den der opretter gruppen som teamleder ***
        strSQLt = "INSERT INTO progrupperelationer (projektgruppeid, medarbejderid, teamleder) VALUES ("& lastId &", "& session("mid") &", 1)"
        oConn.execute(strSQLt)
        end if

		
		Response.redirect "projektgrupper.asp?menu=job&lastid="&lastid
		end if


 
    case "red", "opret"
    '*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"

     intOrgVir = 0
	
	else
	strSQL = "SELECT navn, editor, dato, opengp, orgvir FROM projektgrupper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
    strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intOpenGp = oRec("opengp")
    intOrgVir = oRec("orgvir")

	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
 %>   






<!--redigere-->
<div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Projektgrupper (rediger)</u>
        </h3>

        <form action="projektgrupper.asp?menu=job&func=<%=dbfunc%>" method="post">
	    <input type="hidden" name="id" value="<%=id%>">

        <div class="row">
        <div class="col-lg-10">&nbsp</div>
        <div class="col-lg-2 pad-b10"><button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button></div>
        
        
        </div>



          <div class="portlet-body">
         <section>
             <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Stamdata</h4>
                        </div>
                    </div>

             <div class="row">
                 <div class="col-lg-1 pad-t10">&nbsp</div>
                 <div class="col-lg-2 pad-t10">Navn:&nbsp<span style="color:red;">*</span></div>
                 <div class="col-lg-3 pad-t10">   
                  <input name="FM_navn" type="text" class="form-control input-small" value="<%=strNavn %>" placeholder="Navn"/>
                 </div>
             </div>
                 
                 
                <%if cint(intOrgVir) = 1 then
	            intOrgVirCHK1 = "CHECKED"
                intOrgVirCHK0 = ""
	            else
	            intOrgVirCHK1 = ""
                intOrgVirCHK0 = "CHECKED"
	            end if %>

                 <div class="row">
                     <div class="col-lg-1">&nbsp</div>
                     <div class="col-lg-4 pad-t20"><input name="FM_orgvir" value="1" type="radio" <%=intOrgVirCHK1 %>/>&nbsp;<b>Organisatorisk</b> eller&nbsp;&nbsp;<input name="FM_orgvir" value="0" type="radio" <%=intOrgVirCHK0 %>/>&nbsp;<b>Virtuel</b> gruppe.<br /><br />
                         <div style="background-color:#CCCCCC; padding:10px;">En medarbejder kan kun være medlem af <u>en</u> organisatorisk gruppe, mens virtullegrupper kan bruges til projekter, eller indele medarbejdere efter kompetencer.</div>
                     </div>
                 </div>


                 <%if intOpenGp = "1" then
	            intOpenGpCHK = "CHECKED"
	            else
	            intOpenGpCHK = ""
	            end if %>

                  <div class="row">
                     <div class="col-lg-1">&nbsp</div>
                     <div class="col-lg-4 pad-t20"><input id="Checkbox1" name="FM_opengp" value="1" type="checkbox" <%=intOpenGpCHK %>/>&nbsp;Åben gruppe (medarbejdere må selv tilføje sig - kun virtuelle)</div>
                 </div>

                 <% if func = "opret" then %>
                 <div class="row">
                 <div class="col-lg-1">&nbsp</div>
                 <div class="col-lg-4"><input id="Checkbox2" name="FM_tilfoj_m" value="1" type="checkbox"  CHECKED/>&nbsp;Tilføj dig selv som teamleder i gruppen</div>
                 </div>
                 <% end if %>

           </div>
         </section>
               <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
        </div>
        </form>
       </div>

</div>
<%
case else    


        if len(trim(request("lastid"))) <> 0 then
        lastid = request("lastid")
        else
        lastid = 0
        end if

%>



<script src="js/projektgrupper.js" type="text/javascript"></script>
<div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Projektgrupper</u>
        </h3>

          <form action="projektgrupper.asp?menu=job&func=opret&id=0" method="post">
              <input type="hidden" name="lto" id="lto" value="<%=lto%>">
            <section>
              
                    
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                          <button class="btn btn-sm btn-success pull-right"><b>Opret ny +</b></button><br />&nbsp;
                            </div>
                        </div>
            </section>
         </form>

          <div class="portlet-body">

              <table id="progrp_list" class="table dataTable table-striped table-bordered table-hover ui-datatable">                 
               <thead>
                   <tr>
                       <th style="width: 5%">Id</th>
                       <th style="width: 55%">Projektgruppe</th>
                       <th style="width: 20%">Medlemmer</th>
                       <th style="width: 8%">Orga./Virtuel?</th>
                       <th style="width: 7%">Privat/åben?</th>
                       <th style="width: 5%">Slet</th>
                   </tr>
               </thead>
               <tbody>
                   <%
	                    strSQL = "SELECT id, navn, opengp, orgvir FROM projektgrupper WHERE id > 1 ORDER BY navn"
	                    oRec.open strSQL, oConn, 3
	
	                   
	                    n = 0
	                    while not oRec.EOF


                            
				                
                            
				                
	                         
                                call fTeamleder(session("mid"), oRec("id"))

                        if cint(lastid) = cint(oRec("id")) then
                       trBgCol = "#FFFFE1" 
                       else
                       trBgCol = ""
                       end if
                       
                      
                       antalMediPgrpX = 0
                       call antalMediPgrp(oRec("id"))
				     
                       
                       %>

                        <tr style="background-color:<%=trBgCol%>;">

                  
                       <td><%=oRec("id")%></td>
                       <td>
                           <%if oRec("id") <> "10" AND (oRec("opengp") = 1 OR level = 1 OR erTeamleder = 1) then%>
                           <a href="projektgrupper.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a>
                           <%else %>
                           <%=oRec("navn") %>
                           <%end if %>
                       </td>
                       <td>
                            <%if oRec("id") <> "10" AND (oRec("opengp") = 1 OR level = 1 OR erTeamleder = 1) then%>
                           <a href="projektgrupper.asp?menu=job&func=med&id=<%=oRec("id")%>">Medlemmer (<%=antalMediPgrpX%>) </a>
                           <%else %>
                           Medlemmer (<%=antalMediPgrpX %>)
                           <%end if %>

                       </td>
                            <td>
                                <%if cint(oRec("orgvir")) = 1 then 'ORGANI
                                 orgVir = "Organisatorisk"   
                                 else
                                 orgVir = "Virtuel"   
                                 end if %>

                                <%=orgVir %>

                            </td>
                       <td>
                        <%if oRec("opengp") = 0 then %>
	                    Privat
	                    <%else %>
	                    Åben
	                    <%end if %>
                       </td>
                       <td style="text-align:center;">
                           <%if oRec("id") <> "10" AND (level = 1 OR erTeamleder = 1) then%>
                           <a href="projektgrupper.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a>
                           <%else %>
                           &nbsp;
                           <%end if %>
                       </td>

                   </tr>
                <%

                    
	                n = n + 1
                    oRec.movenext
                    wend
                    oRec.close
                %>
               </tbody>    
                 <tfoot>
                      <tr>
                       <th>Id</th>
                       <th>Projektgruppe</th>
                       <th>Medlemmer</th>
                       <th>Orga./Virtuel?</th>
                       <th>Privat/åben?</th>
                       <th>Slet</th>
                   </tr>

                 </tfoot>
                   
           </table>

          </div>

      </div>

</div>


<%end select  %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->