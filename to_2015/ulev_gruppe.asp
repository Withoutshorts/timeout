
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->



<%call menu_2014 %>
<div class="wrapper">
    <div class="content">


<%

if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)
        response.end
	end if
	
	func = request("func")
	if len(trim(request("id"))) <> 0 then
	id = request("id")
	else
	id = 0
	end if
    if len(trim(request("lastid"))) <> 0 then
    lastid = request("lastid")
    else
    lastid = 0
    end if
        'end if

		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function

select case func 
case "slet"
        oskrift = "Underleverandør grp. / Salgsomkostninger" 
        slttxta = "Du er ved at <b>SLETTE</b> en Underlev. / Salgsomkostings gruppe. Er dette korrekt?"
        slttxtb = "" 
        slturl = "ulev_gruppe.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)


case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM job_ulev_jugrp WHERE jugrp_id = "& id &"")
	Response.redirect "ulev_gruppe.asp?menu=job&shokselector=1"
	

    case "opdliste"


    '**** Tilknytter / Sletter / Opdaterer Uleverandører ***'
	ju_favorit = id				
							
							
						For u = 1 to 20
							    
                            ulevid = request("ulevid_"&u&"")

                            if len(trim(request("ulevfravalgt_"&u&""))) <> 0 then
                            ulevfravalgt = 1
                            else
                            ulevfravalgt = 0
                            end if

							ulevnavn = replace(request("ulevnavn_"&u&""), "'", "''")
							    
							ulevpris = replace(request("ulevpris_"&u&""), ".", "")
							ulevpris = replace(ulevpris, ",", ".")

                            ulevstk = replace(request("ulevstk_"&u&""), ".", "")
							ulevstk = replace(ulevstk, ",", ".")

                            ulevstkpris = replace(request("ulevstkpris_"&u&""), ".", "")
							ulevstkpris = replace(ulevstkpris, ",", ".")
							    
							ulevfaktor = replace(request("ulevfaktor_"&u&""), ".", "")
							ulevfaktor = replace(ulevfaktor, ",", ".")
							    
							ulevbelob = replace(request("ulevbelob_"&u&""), ".", "")
							ulevbelob = replace(ulevbelob, ",", ".")
    							
    					    ju_fase = "" 'replace(request("ulevfase_"&u&""), "'", "''")
    							
		                    if len(trim(ulevnavn)) <> 0 then
					            
					            if len(trim(ulevpris)) <> 0 then
    							ulevpris = ulevpris
    							    
    							    call erDetInt(ulevpris)
    							    if isInt = 0 then
    							    ulevpris = ulevpris
    							    else
    							    ulevpris = 0
    							    end if
    							    isInt = 0
    							    
    							else
    							ulevpris = 0
    							end if
    							

                                if len(trim(ulevstk)) <> 0 then
    							ulevstk = ulevstk
    							    
    							    call erDetInt(ulevstk)
    							    if isInt = 0 then
    							    ulevstk = ulevstk
    							    else
    							    ulevstk = 0
    							    end if
    							    isInt = 0
    							    
    							else
    							ulevstk = 0
    							end if


                                if len(trim(ulevstkpris)) <> 0 then
    							ulevstkpris = ulevstkpris
    							    
    							    call erDetInt(ulevstkpris)
    							    if isInt = 0 then
    							    ulevstkpris = ulevstkpris
    							    else
    							    ulevstkpris = 0
    							    end if
    							    isInt = 0
    							    
    							else
    							ulevstkpris = 0
    							end if

    							
    							if len(trim(ulevfaktor)) <> 0 then
    							ulevfaktor = ulevfaktor
    							    
    							    call erDetInt(ulevfaktor)
    							    if isInt = 0 then
    							    ulevfaktor = ulevfaktor
    							    else
    							    ulevfaktor = 0
    							    end if
    							    isInt = 0
    							    
    							else
    							ulevfaktor = 0
    							end if
    							
    							if len(trim(ulevbelob)) <> 0 then
    							ulevbelob = ulevbelob
    							        
    							    call erDetInt(ulevbelob)
    							    if isInt = 0 then
    							    ulevbelob = ulevbelob
    							    else
    							    ulevbelob = 0
    							    end if
    							    isInt = 0
    							        
    							else
    							ulevbelob = 0
    							end if
    							
                                

    							
						        
                                       if cint(ulevid) <> 0 then
                                

                                            strSQLyuLevOpd = "UPDATE job_ulev_ju SET "_
                                            &" ju_navn = '"& ulevnavn &"', ju_ipris = "& ulevpris &", "_
							                &" ju_faktor = "& ulevfaktor &", ju_belob = "& ulevbelob &", ju_fase = '"& ju_fase &"', ju_stk = "& ulevstk &", ju_stkpris = "& ulevstkpris &", ju_fravalgt = "& ulevfravalgt &" WHERE ju_id ="& ulevid

                                            'Response.Write strSQLyuLevOpd & "<br>"
                                            'Response.flush
                                            oConn.execute(strSQLyuLevOpd)

                                        else
						                    strSQLInsUlev = "INSERT INTO job_ulev_ju (ju_navn, ju_ipris, ju_faktor, ju_belob, ju_fase, ju_favorit, ju_stk, ju_stkpris, ju_fravalgt) VALUES "_
							                &" ('"& ulevnavn &"', "& ulevpris &", "_
							                &""& ulevfaktor &", "& ulevbelob &", '"& ju_fase &"', "& ju_favorit &", "& ulevstk &", "& ulevstkpris &", "& ulevfravalgt &")"  
							        
                                            oConn.execute(strSQLInsUlev)
                                            'Response.Write strSQLInsUlev & "<br>"
                                            'Response.flush
                                
                                        end if

                                else

                                    
                                    if cint(ulevid) <> 0 then
                                    strSQLDelUlev = "DELETE FROM job_ulev_ju WHERE ju_id = "& ulevid
                                    'Response.Write strSQLDelUlev & "<br>"
                                    'Response.flush
							        oConn.execute(strSQLDelUlev)
                                    end if

                               
    							
							    end if 'ulevnavn
    							
							  
							
							next


                            'Response.end
                            response.redirect "ulev_gruppe.asp?lastid="&ju_favorit


	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
		
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		forvalgt = request("FM_forvalgt")

        '** Nulstiller forvalgt **'
        oConn.execute("UPDATE job_ulev_jugrp SET jugrp_forvalgt = 0 WHERE jugrp_id <> 0 ")
		
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO job_ulev_jugrp (jugrp_navn, jugrp_editor, jugrp_dato, jugrp_forvalgt) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& forvalgt &")")
		else
		oConn.execute("UPDATE job_ulev_jugrp SET jugrp_navn ='"& strNavn &"', jugrp_editor = '" &strEditor &"', jugrp_dato = '" & strDato &"', jugrp_forvalgt = "& forvalgt &" WHERE jugrp_id = "&id&"")
		end if
		
		Response.redirect "ulev_gruppe.asp?menu=job&shokselector=1"
		end if
	
	
	case "redgrp"

    %>
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
            <script  src="inc/jugrp_jav.js"></script>

    <%
    strSQLgrp = "SELECT jugrp_navn FROM job_ulev_jugrp WHERE jugrp_id = "& id
    oRec.open strSQLgrp, oConn, 3
    if not oRec.EOF then
    grpNavn = oRec("jugrp_navn")
    
    %>
<!--Redigering af gruppe-->
   <div class="container">
   <div class="portlet">
    <h3 class="portlet-title"><u>Ulev. / Salgsomkost. grp: <%=oRec("jugrp_navn")%> </u></h3>
    
    <form action="ulev_gruppe.asp?func=opdliste&id=<%=id %>" method="post">
        <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
            </div>
        </div>

    <div class="portlet-body">
        
   
    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
 	<thead>
       <tr>

		<th ><b>* Navn</b</th>
        <th ><b>Antal Stk.</b></th>
        <th ><b>Stk. pris</b></th>
		<th ><b>Indkøbspris</b></th>
		<th ><b>Faktor</b></th>
        <th ><b>Salgspris <%=basisValISO %></b></th>
        <th ><b>Fravalgt</b></th>
      
		
    </tr>
	</thead>
        <tbody>
      <%
	
                         
        '*** Ulev ***'
	    dim u_navn, u_ipris, u_faktor, u_belob, u_fase, u_id, u_istk, u_istkpris, u_fravalgt
	    redim u_navn(20), u_ipris(20), u_faktor(20), u_belob(20), u_fase(20), u_id(20), u_istk(20), u_istkpris(20), u_fravalgt(20)
		                
		        strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, ju_id, ju_fase, ju_stk, ju_stkpris, "_
		        &" ju_belob, ju_fravalgt FROM job_ulev_ju WHERE ju_favorit = "& id  & " AND ju_jobid = 0 ORDER BY ju_fase, ju_navn "
		        oRec2.open strSQLUlev, oConn, 3
		                        
		        u = 1
		                        
		        while not oRec2.EOF 
		                            

                    u_fravalgt(u) = oRec2("ju_fravalgt")
		            u_navn(u) = oRec2("ju_navn")
		            u_id(u) = oRec2("ju_id")
		                     
		            u_ipris(u) = oRec2("ju_ipris")
		            u_faktor(u) = oRec2("ju_faktor")
		            u_belob(u) = oRec2("ju_belob")
		            u_id(u) = oRec2("ju_id")
                    u_fase(u) = "" 'oRec2("ju_fase")

                    u_istk(u) = oRec2("ju_stk")
                    u_istkpris(u) = oRec2("ju_stkpris")
		                            
		            u = u + 1
		                            
		            oRec2.movenext
		            wend
		            oRec2.close
		                         
		                       
		                 
		               
                        
                         
                      
		                
	    for u = 1 to 20 
		                
	    select case right(u, 1)
	    case 0,2,4,6,8
	    bgulev = "#FFFFFF"
	    case else
	    bgulev = "#EFF3FF"
	    end select
		                
		    uvzb = "visible"
		    udsp = ""
		                
		    if len(trim(u_id(u))) <> 0 then
            u_id(u) = u_id(u)
            else
            u_id(u) = 0
            end if

            if u_fravalgt(u) <> 0 then
            fraVChk ="CHECKED"
            else
            fraVChk =""
            end if
		                
	    %>
		                
		                
		    <tr bgcolor="<%=bgulev %>" id="ulevlinie_<%=u%>" style="visibility:<%=uvzb%>; display:<%=udsp%>;">
                         
		    <input id="Hidden1" name="ulevid_<%=u%>" value="<%=u_id(u) %>" type="hidden" />
            <input type="hidden" id="ulevfase_<%=u%>" name="ulevfase_<%=u%>" value="<%=u_fase(u) %>">
            <td><input type="text" id="ulevnavn_<%=u%>" name="ulevnavn_<%=u%>" value="<%=u_navn(u) %>" style="width:250px;" class="form-control input-small"></td>
            <td><input class="beregn_tf form-control input-small" type="text" id="ulevstk_<%=u%>" name="ulevstk_<%=u%>" value="<%=replace(formatnumber(u_istk(u), 2), ".", "") %>" style="width:95px;"></td>
            <td><input class="beregn_tf form-control input-small" type="text" id="ulevstp_<%=u%>" name="ulevstkpris_<%=u%>" value="<%=replace(formatnumber(u_istkpris(u), 2), ".", "") %>" style="width:95px;"></td>
		    <td><input class="beregn_tf form-control input-small" type="text" id="ulevpri_<%=u%>" name="ulevpris_<%=u%>" value="<%=replace(formatnumber(u_ipris(u), 2), ".", "") %>" style="width:95px;"></td>
		    <td><input class="beregn_tf form-control input-small" type="text" id="ulevfak_<%=u%>" name="ulevfaktor_<%=u%>" value="<%=replace(formatnumber(u_faktor(u), 2), ".", "") %>" style="width:40px;"></td>
		    <td><input class="beregn_to form-control input-small" type="text" id="ulevbel_<%=u%>" name="ulevbelob_<%=u%>" value="<%=replace(formatnumber(u_belob(u), 2), ".", "") %>" style="width:95px;"></td>
            <td><input type="checkbox" name="ulevfravalgt_<%=u%>" id="ulevfravalgt_<%=u%>" value="1" <%=fraVChk %> /></td>
		                    
		                    
	    </tr>
		                
	    <%next %>
        </tbody>
    </table>
    </div>
	</form>

    </div>
    </div>

    <%
        end if
    oRec.close
    %>

<%

	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intForvalgt = 0
	
	else
	strSQL = "SELECT jugrp_navn, jugrp_forvalgt, jugrp_editor, jugrp_dato FROM job_ulev_jugrp WHERE jugrp_id=" & id
	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("jugrp_navn")
	strDato = oRec("jugrp_dato")
	strEditor = oRec("jugrp_editor")
	intForvalgt = oRec("jugrp_forvalgt")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--Redigering af Underleverandør grp. / Salgsomkostninger-->	
    
            <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Underleverandør grp. / Salgsomkostninger</u></h3>

                <form action="ulev_gruppe.asp?menu=crm&func=<%=dbfunc%>" method="post">
                <input type="hidden" name="id" value="<%=id%>">
                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                        <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                        </div>
                    </div>

                <div class="well well-white">
                    <div class="portlet-body">
                        <div class="row">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-2">Gruppenavn:</div>
                            <div class="col-lg-3"><input class="form-control input-small" name="FM_navn" value="<%=strNavn%>"></div>
                            <div class="col-lg-2"></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-2">Forvalgt<br /> (ved joboprettelse):</div>
                            <div class="col-lg-1">
                                 <%if cint(intForvalgt) <> 0 then
                                    fvCHK1 = "CHECKED"
                                    fvCHK0 = ""
                                    else
                                    fvCHK0 = "CHECKED"
                                    fvCHK1 = ""
                                    end if %>

                                    <input id="Radio1" type="radio" name="FM_forvalgt" value="0" <%=fvCHK0 %> /> Nej
                                    <br />
                                    <input id="Radio2" type="radio" name="FM_forvalgt" value="1" <%=fvCHK1 %> /> Ja
                            </div>
                        </div>
                    </div>
                </div>
                <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
             </form> 
           </div>
           </div>

<%case else %>
<!--Liste-->
<script src="js/ulev_gruppeliste.js" type="text/javascript"></script>
       <div class="container">
        <div class="portlet">
            <h3 class="portlet-title">Underleverandør grp. / Salgsomkostninger</h3>
        

            <form action="ulev_gruppe.asp?menu=job&func=opret" method="post">
            <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                                <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                            </div>
                 </div>
              </section>
         </form>

        <div class="portlet-body">

            <table id="ulevgrpliste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                <thead>
                    <tr>
                        <th style="width: 2%">id</th>
                        <th style="width: 15%">Underleverandør gruppe / Salgsomkostninger</th>
                        <th style="width: 15%">Se gruppe (antal emner i grp.)</th>
                        <th style="width: 5%">Slet</th>
                    </tr>
                </thead>
                <tbody>
                    <%
	                    strSQL = "SELECT jugrp_id, jugrp_navn, jugrp_forvalgt FROM job_ulev_jugrp WHERE jugrp_id <> 0 ORDER BY jugrp_navn"
	                    x = 0
	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF 


                        if cint(lastid) = oRec("jugrp_id") then
                            bgcol = "#FFFF99"

                        else
        
                            select case right(x, 1)
                            case 0,2,4,6,8
                            bgcol = "#FFFFFF"
                            case else
                            bgcol = "#EFf3FF"
                            end select

                        end if

	                %>
                    <tr>
                        <td><%=oRec("jugrp_id")%></td>
                        <%
					        '** Henter aktiviteter i den aktuelle gruppe ****
                            intAntal = 0
					        strSQL2 = "SELECT count(ju_id) AS antal FROM job_ulev_ju WHERE ju_favorit = "&oRec("jugrp_id")&" AND ju_jobid = 0"
					        oRec2.open strSQL2, oConn, 3
					        if not oRec2.EOF then
					        intAntal = oRec2("antal")
					        end if
					        oRec2.close 
					    %>
                        <td><a href="ulev_gruppe.asp?menu=job&func=red&id=<%=oRec("jugrp_id")%>"><%=oRec("jugrp_navn")%> </a>

                            <%if oRec("jugrp_forvalgt") = 1 then %>
                            &nbsp;(forvalgt)
                            <%end if %>

                        </td>
                        <td><a href='ulev_gruppe.asp?menu=job&func=redgrp&id=<%=oRec("jugrp_id")%>&stamakgrpnavn=<%=oRec("jugrp_navn")%>' class='vmenuglobal'>Se / Rediger poster i grp.&nbsp;</a>(<%=intAntal%>)</td>
                        <%if cint(intAntal) = 1 then %>
			            <td>&nbsp</td>   
                        <%else %>
                        <td style="text-align:left;"><a href="ulev_gruppe.asp?menu=job&func=slet&id=<%=oRec("jugrp_id")%>"><span style="color:darkred; display: block; text-align: left;" class="fa fa-times"></span></a></td>
                        <%end if %>
                    </tr>
                    <%
	                    x = x + 1
	                    oRec.movenext
	                    wend
	                %>
                </tbody>
                <tfoot>
                    <tr>
                        <th>Gruppenavn</th>
                        <th>Underleverandør grp. / Salgsomkostninger</th>
                        <th>Se gruppe (antal emner i grp.)</th>
                        <th>Slet</th>
                    </tr>
                </tfoot>
            </table>

        </div>
        </div>
       </div>
    


<%end select %>
 
</div> <!--content-->
</div> <!--wrapper-->

<!--#include file="../inc/regular/footer_inc.asp"-->