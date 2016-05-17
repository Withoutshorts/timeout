

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#indclude file="medarbtyper.asp"-->

    
<!--#include file="../timereg/inc/timereg_dage_2006_inc.asp"-->

<script type="text/javascript" src="js/libs/excanvas.compiled.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.orderBars.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.pie.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.resize.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.stack.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.tooltip.min.js"></script>
<script type="text/javascript" src="js/demos/flot/stacked-vertical_ugetotal.js"></script>
<script type="text/javascript" src="js/demos/flot/donut.js"></script>
<script type="text/javascript" src="js/demos/flot/vertical.js"></script>
<script type="text/javascript" src="js/demos/flot/scatter.js"></script>
<script type="text/javascript" src="js/demos/flot/line.js"></script>
<script type="text/javascript" src="js/demos/flot/horizontal-ugeseddelmånedtotal.js"></script>
<link rel="stylesheet" href="../../bower_components/bootstrap-datepicker/css/datepicker3.css">



<%call menu_2014 %>

<div class="wrapper">
    <div class="content">

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
                pa = 0
                if len(trim(request("jq_pa") )) <> 0 then
                pa = request("jq_pa") 
                else
                pa = 0
                end if
        
                'pa = 0
            

                if filterVal <> 0 then
            
                 
    
                strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                         

               if pa = "1" then
               strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               &" WHERE tu.medarb = "& usemrn &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"   


               else


                        '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 


           




               strSQL= "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"      
    

               
            
                end if

                 'response.write "strSQL " &strSQL
                 'response.end


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


                
                
                if cint(showAkt) = 1 then 
                 
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
 

if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
	response.End
    end if
	
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "ugeseddel_2011.asp"
	media = request("media")




    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> 0 then
    nomenu = 1
    else
    nomenu = 0
    end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if
    


            
         

    level = session("rettigheder")
 
    rdir = request("rdir")
    select case rdir
    case "logindhist"
    rdirfile = "logindhist_2011.asp"
    case "godkenduge"
    rdirfile = "godkenduge.asp"
    case else
	rdirfile = "ugeseddel_2011.asp"
	end select
    

    if len(trim(request("medarbsel_form"))) <> 0 then

        'response.write "HER"
        'response.end

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



    
    call ersmileyaktiv()
	

    call medarb_teamlederfor
    
        
    
        
    varTjDatoUS_man = request("varTjDatoUS_man")

    'if len(trim(request("varTjDatoUS_son"))) <> 0 then
	'varTjDatoUS_son = request("varTjDatoUS_son")
    'else
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'end if

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





    call positiv_aktivering_akt_fn()
    call smileyAfslutSettings()

	'Response.Write "media " & media
	
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	
    select case func 
	case "-"

    case "opdaterstatus"


        
        'response.write "request(ids)" & request("ids")
        
          ujid = split(request("ids"), ",")

                if len(trim(request("FM_godkendt"))) <> 0 then
                uGodkendt = request("FM_godkendt")
                else
                uGodkendt = 0
                end if
	
    for u = 0 to UBOUND(ujid)
	
	            editor = session("user")


                'if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                'uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                'else
                'uGodkendt = 0
                'end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u)
				
			    'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

       'response.end

    'Response.Redirect "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
	case "godkendugeseddel"
	
	'*** Godkender ugeseddel ***'
	
	    varTjDatoUS_manSQL = varTjDatoUS_man
        varTjDatoUS_sonSQL = year(varTjDatoUS_son)&"/"&month(varTjDatoUS_son)&"/"&day(varTjDatoUS_son)
	    
	    strSQLup = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tmnr = "& usemrn 
	    if cint(SmiWeekOrMonth) = 0 then
        strSQLup = strSQLup & " AND tdato BETWEEN '"& varTjDatoUS_manSQL &"' AND '" & varTjDatoUS_sonSQL & "'" 
        else
        varTjDatoUS_man_mth = datepart("m", varTjDatoUS_man,2,2)
        strSQLup = strSQLup & " AND MONTH(tdato) = '"& varTjDatoUS_man_mth & "'" 
        end if

        strSQLup = strSQLup & " AND godkendtstatus <> 1" 

	   ' strSQLup = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tmnr = "& usemrn & " AND tdato BETWEEN '"& varTjDatoUS_man &"' AND '" & varTjDatoUS_son & "' AND godkendtstatus <> 1" 
	    oConn.execute(strSQLup)
	    
        'if session("mid") = 1 then
	    'Response.Write strSQLup
	    'Response.end
        'end if
        
        
        '*** Godkend uge status ****'
        call godekendugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man)
	    
	
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "afvisugeseddel"
	
	'*** Afviser ugeseddel ***'
    if len(trim(request("FM_afvis_grund"))) <> 0 then
    txt = replace(request("FM_afvis_grund"), "'", "")
    else
    txt = ""
    end if

	call afviseugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)
	    
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
    
    
    case "adviserugeafslutning"

    call afslutugereminder(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)

  
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "slet_tip"

        

        id = request("id")

        slttxt = "Du er ved at slette en timeregistrering uden match.<br> Er Du sikker på Du vil gøre dette?"
        slturl = rdirfile & "?func=slet_tip_ok&id="&id&"&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

       call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)

    

    case "slet_tip_ok"

        id = request("id")

        sqlInTimerImpTemp = "DELETE FROM timer_import_temp WHERE id = "& id
        oConn.Execute(sqlInTimerImpTemp)

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

    case else

   

    call akttyper2009(2)
	
	
	
	
	%>
    <SCRIPT language=javascript src="inc/ugeseddel_2011_jav.js"></script>    

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ugeseddel:</u></h3>
                <div class="portlet-body">
                    
                    <div class="row">
                        <form id="filterkri" method="post" action="ugeseddel.asp">
                            <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
                            <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
                            <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
                            <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">


                            <%
                                function dageDatoer(vis)

                                    Redim writedagnavn(7)
                                    writedagnavn(7) = tsa_txt_127 '"Sø"
                                    writedagnavn(1) = tsa_txt_128 '"Ma"
                                    writedagnavn(2) = tsa_txt_129 '"Ti"
                                    writedagnavn(3) = tsa_txt_130 '"On"
                                    writedagnavn(4) = tsa_txt_131 '"To"
                                    writedagnavn(5) = tsa_txt_132 '"Fr"
                                    writedagnavn(6) = tsa_txt_133 '"Lø"


                                    '*** Medarb ***'
                                    call meStamdata(usemrn)
                                    call jq_format(meTxt)
                                    meTxt = jq_formatTxt


                                    if vis = 2 then
                                    cspsThis = 5
                                    bdsThis = 1
                                    else
                                    cspsThis = 4
                                    bdsThis = 1
                                    end if

                                    dageDatoerTxt = "<tr><td colspan=2 bgcolor=""#FFFFFF"" style=""padding:5px 5px 4px 5px; border-bottom:"&bdsThis&"px #cccccc solid;""><b>"&tsa_txt_005&": "& datepart("ww", tjekdag(1), 2,2) & "&nbsp;&nbsp;&nbsp;" & datepart("yyyy", tjekdag(1), 2,2) &"</b></td>"
                                    dageDatoerTxt = dageDatoerTxt & "<td colspan="&cspsThis+3&" bgcolor=""#FFFFFF"" align=right style=""padding:5px 5px 4px 5px; border-bottom:"&bdsThis&"px #cccccc solid; font-size:10px;""><i> "& meTxt &", "& normTimerprUge &" t. pr. uge</i></td>"

                                    dageDatoerTxt = dageDatoerTxt & "</tr>"


                                    dageDatoerTxt = dageDatoerTxt & "<tr><td colspan=2 bgcolor=""#ffffff"" valign=top style=""width:290px; border-bottom:"&bdsThis&"px #8caae6 solid; border-right:"&bdsThis&"px #8caae6 solid; white-space:nowrap;"">"

                                    if vis = 1 then
    
                                        dageDatoerTxt = dageDatoerTxt & "<img src=""../ill/nut_and_bolt.png"" border=""0""><b>"& tsa_txt_125 &"</b>"

    
                                        select case lto
                                        case "unik", "mmmi"
                                        case else

                                        if visSimpelAktLinje <> "1" AND visHRliste <> "1" then
                                        dageDatoerTxt = dageDatoerTxt & "<br /><span style=""font-size:9px;"">Periode | "& tsa_txt_222 & " ell. "& tsa_txt_186 &" | Forkalk. timer | Real. timer" 
    
                                        if (cint(aktBudgettjkOn) = 1) then
                                        dageDatoerTxt = dageDatoerTxt & "<br>| Forecast"
                                        end if

                                         dageDatoerTxt = dageDatoerTxt & "</span>"

                                        end if

                                        end select


                                        dageDatoerTxt = dageDatoerTxt & "</td>"

                                    else
                                    dageDatoerTxt = dageDatoerTxt & "&nbsp;</td>"
                                    end if

                                    'dageDatoerTxt = dageDatoerTxt & "<td bgcolor=""#ffffff"" style=""border-bottom:"&bdsThis&"px #8caae6 solid; border-right:"&bdsThis&"px #8caae6 solid;"">&nbsp;</td>"




                                    t = 1
                                    for t = 1 to 7
	                                    '*** Gennemløber de 7 dage ***
	
	
	                                    call helligdage(tjekdag(t), 0, lto)
	
	
		                                    if (t = 6  OR t = 7) OR erHellig = 1 then
		                                    borderColor = "#999999"
		                                    border = bdsThis
		                                    bgcoldg = "gainsboro"
		                                    else
		                                    borderColor = "#8caae6"
		                                    border = bdsThis
		                                    bgcoldg = "#FFFFFF"
		                                    end if
    
                                        if formatDatetime(now, 2) = formatDatetime(tjekdag(t), 2) then
	                                    borderColor = "red" '"#ffff99"
	                                    border = 2
                                        end if

    
                                        '*** ÆØÅ **'
                                        call jq_format(writedagnavn(t))
                                        writedagnavn(t) = jq_formatTxt

	                                    dageDatoerTxt = dageDatoerTxt & "<td valign=top style='width:65px; background-color: "&bgcoldg&"; border-right:"&bdsThis&"px #8caae6 solid;  border-bottom:"&border&"px "&borderColor&" solid;' class='lille'>"& writedagnavn(t) &"<br>"
	                                    dageDatoerTxt = dageDatoerTxt & formatdatetime(tjekdag(t), 2) &"<br>"
	
                                        'dageDatoerTxt = dageDatoerTxt & "<td>" & formatdatetime(tjekdag(t), 2) 
                                        call jq_format(helligdagnavnTxt)
                                        helligdagnavnTxt = jq_formatTxt

                                        'call helligdage(tjekdag(t), 0)
                                        dageDatoerTxt = dageDatoerTxt & "<span style=""font-size:9px; color:#5c75AA; line-height:9px;"">"& helligdagnavnTxt & "</span></td>"

                                    next


                                    if vis = 2 then
                                    dageDatoerTxt = dageDatoerTxt & "<td bgcolor=#ffffff class=lille valign=top align=right style=""border-bottom:"&bdsThis&"px #8caae6 solid; border-right:"&bdsThis&"px #8caae6 solid;""><br>Total</td>"
                                    end if

                                    dageDatoerTxt = dageDatoerTxt & "</tr>"

                                    end function



                            %>


                            <%   
                                

                         call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours)

                                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                                manTimer = replace(manTimer, ",",".")
                                tirTimer = replace(tirTimer, ",",".")
                                onsTimer = replace(onsTimer, ",",".")
                                torTimer = replace(torTimer, ",",".")
                                lorTimer = replace(lorTimer, ",",".")
                                sonTimer = replace(sonTimer, ",",".")

                             call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)

                            norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

                            norm_ugetotal = 34
                            weeknormpro = ugetotal / norm_ugetotal * 100
                            weeknormpro = replace(weeknormpro, ",",".")

                            %>
                            

                            <input type="hidden" id="dagman" value="<%=manTimer %>" />
                            <input type="hidden" id="dagtir" value="<%=tirTimer %>" />
                            <input type="hidden" id="dagons" value="<%=onsTimer %>" />
                            <input type="hidden" id="dagtor" value="<%=torTimer %>" />
                            <input type="hidden" id="dagfre" value="<%=freTimer %>" />
                            <input type="hidden" id="daglor" value="<%=lorTimer %>" />
                            <input type="hidden" id="dagson" value="<%=sonTimer %>" />
                            <%if session("mid") <> 1 then %>
                            
     
                            <div class="col-lg-3">
                                <input type="checkbox" onchange="this.form.submit()" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %>/>
                                <%call medarb_vaelgandre %>
                                <br /><br /><br />
                                <div class="col-lg-8">reg. timer denne uge:</div>
                                <div class="col-lg-1"><h4><%=ugetotal %></h4></div>
                                <br />
                                <table class="table table">
                                    <tr>
                                        <th colspan="11"><div style="background-color:lightgray; height:30px; width: 100%;"><div style="background-color:mediumseagreen; height:30px; width:<%=weeknormpro %>%;"></div></div></th>
                                    </tr>
                                    <tr>
                                        <th style="border-top:none; width:10%;"><h6>00</h6></th>
                                        <th style="border-top:none; width:10%;"><h6></h6></th>
                                        <th style="border-top:none; width:10%;"><h6>20</h6></th>
                                        <th style="border-top:none; width:10%;"><h6></h6></th>
                                        <th style="border-top:none; width:10%;"><h6>40</h6></th>
                                        <th style="border-top:none; width:10%;"><h6></h6></th>
                                        <th style="border-top:none; width:10%;"><h6>60</h6></th>
                                        <th style="border-top:none; width:10%;"><h6></h6></th>
                                        <th style="border-top:none; width:10%;"><h6>80</h6></th>
                                        <th style="border-top:none; width:10%;"><h6></h6></th>
                                        <th style="border-top:none; width:10%;"><h6>100</h6></th>
                                        
                                    </tr>
                                    <tr>
                                        <th style="text-align: center; border-top:none;" colspan="11">Norm timer: <%=norm_ugetotal %></th>
                                    </tr>
                                    <tr>
                                        <th style="text-align: center; border-top:none;" colspan="11">Opgjort: <%=weeknormpro %>%</th>
                                    </tr>
                                </table>
                            </div>
                            <div class="col-lg-2">&nbsp</div>
                            <div class="col-lg-5"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>
                            <%else%>
                            
                            <%end if  %>  
            
                        </form>
                        </div>
                 
                    <!--<div class="row">
                            
                            <div class="col-lg-7">&nbsp</div>
                            <div class="col-lg-1" style="background-color:yellowgreen;">
                                <h5 class="row-stat-value"><b><%=ugetotal %> t.</b></h5>
                            </div> 
                           <div class="col-lg-6">&nbsp</div>
                            <div class="col-sm-6 col-md-1">
                                <div class="row-stat">
                                <h4 class="row-stat-value"><%=ugetotal %> t.</h4>
                                </div> 
                            </div> 
                        
                        </div> -->
                        
                        
                   </div>   
                </div>
            </div>

        <%
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
        %>

        </div>
    </div>

<% end select %>
<!--#include file="../inc/regular/footer_inc.asp"-->
