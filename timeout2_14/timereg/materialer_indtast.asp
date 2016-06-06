<%response.buffer = true%>


<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="Mon, 22 Jul 2002 11:12:01 GMT">
    

<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/mat_func.asp"-->


<%

 session.lcid = 1030 'DK
	
'**** Søgekriterier AJAX **'
'section for ajax calls
 if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_sogjob_mat"

        vispasogluk = request("vispasogluk")
        ignprojekgrp = request("ignprojekgrp")

        if len(trim(request("jobsogVal"))) <> 0 then
        jobsogVal = request("jobsogVal")
        else
        jobsogVal = "fdsfs332#43"
        end if

        matreg_jobid = request("matreg_jobid") 


        vis = 0

        'response.write "ignprojekgrp: "& ignprojekgrp & " ignprojekgrp: "& ignprojekgrp & "jobsogVal: " & jobsogVal 

            
        call jobliste(vis, vispasogluk, ignprojekgrp, jobsogVal, matreg_jobid)

        response.end

         case "FN_vis_pas_medarb"

                usemrn = request("matreg_medid")
                vis_passive = request("vispassive_medarb")
                jqHTML = "1"

                
                call selmedarbOptions

                '*** ÆØÅ **'
                call jq_format(strMedarbOptionsHTML)
                strMedarbOptionsHTML = jq_formatTxt

        response.write strMedarbOptionsHTML

        response.end

        case "FN_seneste_mat"

                usemrn = request("matreg_medid")
                
                sogBilagOrJob = 0 'request("sogBilagOrJob")
                matreg_personlig = request("matreg_personlig")
                matreg_visallemed = request("matreg_visallemed")
                sogliste = request("sogliste")

                if len(trim(sogliste)) <> 0 then
                useSog = 1
                else
                useSog = 0
                end if

                call senstematReg(usemrn, useSog, sogBilagOrJob, matreg_personlig, matreg_visallemed, sogliste)

        case "FN_sog_akt"

        matreg_jobid = request("matreg_jobid")    
        status = 1
        vis = 0

        call aktlisteOptions(matreg_jobid, status, vis, 0)

         '*** ÆØÅ **'
        call jq_format(straktOptionlist)
        straktOptionlist = jq_formatTxt

        response.write straktOptionlist
        response.end

        case "FN_sog_lager"

        strGrp = request("matreg_grp")
        strLev = request("matreg_lev")
        strSog = request("matreg_sog")

        'response.write "strLev: "& strLev & "<br>"
        'response.write "strGrp: "& strGrp & "<br>strSog : "& strSog  & "<br>"

       %>
    <form>
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
                <%call matregLagerheader(1) %>
       

            <% 
                
                    useSog = 1

                     if len(trim(strSog)) <> 0 then
                        sogeKri = strSog
                    else
                        sogeKri = ""
                    end if    
                    
                   
                    if len(trim(strLev)) <> 0 AND strLev <> "null" AND strLev <> 0 then
                        sogLevSQLkri = " AND m.leva = "& strLev
                    else
                        sogLevSQLkri = " AND m.leva <> -1 "
                    end if    

                    
                    if len(trim(strGrp)) <> 0 AND strGrp <> "null" AND strGrp <> 0 then
                        sogMatgrpSQLkri = " m.matgrp = "& strGrp
                    else
                        sogMatgrpSQLkri = " m.matgrp <> -1 "
                    end if    
                
                    
                    vis = 1
                
                call matfelter_lager(vis) %>

            </table>
        </form>

<%
     case "FN_indlaes_mat"

       

        '**** Værdier ****'

        strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
        jobid = request("matreg_jobid")
        medid = request("matreg_medid")


        '** Expences / Materialeforbrug
        '** OnTheFly: 1 / Fra lager : 0 
        if len(request("matreg_onthefly")) <> 0 AND request("matreg_onthefly") <> 0 then
		otf = 1 'request("matreg_onthefly")       
		else
		otf = 0
		end if      

        matId = request("matreg_matid")
        aktid = request("matreg_aktid")
        aftid = request("matreg_aftid")

        if len(trim(request("matreg_antal"))) <> 0 then
		intAntal = replace(request("matreg_antal"), ",", ".")
		else
		intAntal = 0
		end if

        'if IsDate(request("matreg_regdato")) = true then
        regdato = request("matreg_regdato")
        'else
        'regdato = year(now)&"/"&month(now)&"/"&day(now)
        'end if
    
        valuta = request("matreg_valuta")

        intkode = request("matreg_intkode")
                
        if len(trim(request("matreg_personlig"))) <> 0 AND request("matreg_personlig") <> 0 then
        personlig = 1
        else
        personlig = 0
        end if
                
                
        if len(request("matreg_bilagsnr")) <> 0 then
        bilagsnr = request("matreg_bilagsnr") 
        else
        bilagsnr = ""
        end if

        pris = request("matreg_pris")
        salgspris = request("matreg_salgspris")
        gruppe = request("matreg_gruppe")

        navn = replace(request("matreg_navn"), "'", "")
       
        if len(trim(request("matreg_varenr"))) <> 0 then
        varenr = request("matreg_varenr")
        else
        varenr = 0
        end if
        
        matreg_opdaterpris = request("matreg_opdaterpris")
        opretlager = request("matreg_opretlager")
        betegnelse = request("matreg_betegn")

        mat_func = request("matreg_func")
        
        'if session("mid") = 1 then
        'response.write "mat_func: "& mat_func &" - pris: "& pris &" - salgspris: "& salgspris &" - onthefly: "& otf &" aktid:"& aktid &" Dato: "& regdato &" Navn: "& navn & " varenr: "& varenr &" opretlager: "& opretlager &" valuta: "& valuta &"<br>"
        'response.end
        'end if
        
        matregid = request("matregid")
        matava = 0
        
        call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava)
	
         

        response.end

end select
response.end
end if
 

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <!--<SCRIPT src="inc/matind_jav.js"></script>-->
    <SCRIPT src="inc/matind_2014_jav.js"></script>

	<%
	
    call smileyAfslutSettings()

	func = request("func")
	
	'Response.Write "func:" & func
	
	'** jobid **'
	if len(trim(request("id"))) <> 0 then
	id = request("id")
	else
	id = 0
	end if

    if len(trim(request("jobid_sel"))) <> 0 then
    id = request("jobid_sel")
    end if
   

    if len(trim(request("mid"))) <> 0 then
	usemrn = request("mid")

    else

	    if len(request.cookies("timereg_2006")("usemrn")) <> 0 then
		usemrn = request.cookies("timereg_2006")("usemrn")
		else
		usemrn = session("mid")
		end if
	
    end if

	if len(request("aftid")) <> 0 then  
	aftid = request("aftid")
	else
	aftid = 0
	end if
	
	if len(trim(request("fromsdsk"))) <> 0 then
	fromsdsk = request("fromsdsk")
	else
	fromsdsk = 0
	end if
	
	if len(request("lastid")) <> 0 then
	lastid = request("lastid")
	else
	lastid = 0
	end if
	
	if len(request("matregid")) <> 0 then
	    matregid = request("matregid")
	    else
	    matregid = 0
	    end if

    
    '***** Vuis passive medarbjedere OTF ****'
    'Response.write "passive:" & request("FM_vis_passive")

    'if len(trim(request("FM_vis_passive"))) <> 0 AND request("FM_vis_passive") <> 0 then
    'vis_passive = 1
    'else
        
    '    if len(trim(request("jobid_sel"))) <> 0 then 'er der søgt
           
    '        vis_passive = 0

    '    else
       
    '     if request.cookies("timereg_2006")("vispassive") <> "" then
    '        vis_passive = request.cookies("timereg_2006")("vispassive")
    '        else
    '        vis_passive = 0
    '    end if

    '    end if
    'end if 

    'response.cookies("timereg_2006")("vispassive") = vis_passive
	
    if cint(vis_passive) = 1 then
    chkPassive = "CHECKED"
    else
    chkPassive = ""
    end if

	'select case lto 
	'case "syncronic"
	'level = 1
	'case else
	level = session("rettigheder")
	'end select
	
	
	
	if len(request("vis")) <> 0 then
	vis = request("vis")
	else
	    
	    '* vis = request.cookies("timereg_2006")("vismatreg")
	    select case lto
	    case "xx"
	    vis = ""
	    case else
	    vis = "otf"
	    end select
	    
	end if

   


	if vis = "otf" then
	Knap1_bg = "#5C75AA"
	Knap2_bg = "#3B5998"
	else
	Knap1_bg = "#3B5998"
	Knap2_bg = "#5C75AA"
	end if
	
	Knap3_bg = "#3B5998"
	Knap4_bg = "#3B5998"
	
	response.cookies("timereg_2006")("vismatreg") = vis
	
	
	'*** LUKKET FOR REDIGERING ***'
	'*** Er Smiley og Auto Gk slået til ? **'
	'*** Auto gk er om periode skal lukkes ved godkend uge **''
	call ersmileyaktiv()
	
	call licensStartDato()
	
	stDato = startDatoAar &"/"& startDatoMd &"/"& startDatoDag
	
	'*** Afsluttede uger for medarb. Finder StDato på Interval ***
	strSQL = "SELECT mf.forbrugsdato FROM materiale_forbrug mf "_
	&" WHERE usrid = "& usemrn &" ORDER BY mf.forbrugsdato, mf.id DESC LIMIT 40" 
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	stDato = year(oRec("forbrugsdato")) &"/"& month(oRec("forbrugsdato")) &"/"& day(oRec("forbrugsdato"))
	end if
	oRec.close
	
	slDato = year(now) &"/"& month(now) &"/"& day(now)
	call afsluger(usemrn, stdato, sldato)
	
	
	
	
	select case func
    case "slet"
    
    medarbId = request("mid")
    
    matid = 0
    matantal = 0
    
    '*** Henter materiale id og antal ***'
    strSQLsel = "SELECT matantal, matid, matnavn FROM materiale_forbrug WHERE id = " & matregid
    oRec.open strSQLsel, oConn, 3
    if not oRec.EOF then
    
    matid = oRec("matid")
    matantal = oRec("matantal")
    matnavn = oRec("matnavn")
    
    end if
    oRec.close
    
    if len(matantal) <> 0 then
    matantal = replace(matantal, ",", ".")
    else
    matantal = 0
    end if

    if len(matid) <> 0 then
    matid = matid
    else
    matid = 0
    end if
    
    '** Opdaterer antal på lager ***
	strSQL2 = "UPDATE materialer SET antal = (antal+("&matantal&")) WHERE id = "& matid
    'Response.Write strSQL2
    'Response.flush
	oCOnn.execute(strSQL2)
	
	'*** Sletter registrering ***
	strSQLdel = "DELETE FROM materiale_forbrug WHERE id = " & matregid
    oConn.execute(strSQLdel) 
    

    '*** Indsætter i delete historik ****'
    matnavn = replace(matnavn, "'", "")
	call insertDelhist("mat", matregid, 0, matnavn, session("mid"), session("user"))

    Response.Redirect "materialer_indtast.asp?id="&id&"&fromsdsk="&fromsdsk&"&aftid="&aftid&"&vis="&vis&"&mid="&medarbId
    	

	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****

		'*** Er flyttet til jquery og cls_mat
	
	
	
	case else
	
	

   
	
	
	strIndlaest = request("FM_indlaest") 
    jobsogval = ""
	
	%>
	
	
	
	
	<div id="sidediv" style="position:absolute; left:90px; top:82px; visibility:visible; width:740px; border:0px #999999 solid;">
	
   
	
	
	<%call matregmenu(0,0,Knap1_bg,Knap2_bg,Knap3_bg,Knap4_bg) 
        
        
     call menu_2014()

	
	
	
	'*** Bilagsnr **'
	bilagsnr = mednr &"-"& year(now) & datepart("m", now, 2,2)
	
	call hentbgrppamedarb(usemrn) %>
	
	<%if vis = "otf" then

        if func = "red" then
	    
	    
	    
	    strSQL = "SELECT matid, matantal, matnavn, matvarenr, matkobspris, matenhed, jobid, matsalgspris, "_
	    &" matgrp, forbrugsdato, valuta, intkode, kurs, bilagsnr, personlig, aktid FROM materiale_forbrug WHERE id = " & matregid
	    
	    'Response.Write "<br><br><br>"& strSQL
	    'Response.flush
	    
	    oRec.open strSQL, oConn, 3
	    if not oRec.EOF then
	    
	    matid = oRec("matid")
	    matantal = oRec("matantal")
	    matnavn = oRec("matnavn")
	    matvarenr = oRec("matvarenr")
	    matkobspris = oRec("matkobspris")
	    matsalgspris = oRec("matsalgspris")
	    matenhed = oRec("matenhed")
	    matgrp = oRec("matgrp")
	    forbrugsdato = oRec("forbrugsdato")
	    
	    valuta = oRec("valuta")
	   
	    intKode = oRec("intkode")
	    id = oRec("jobid") 
	    bilagsnr = oRec("bilagsnr")
	    
	    personlig = oRec("personlig")
        aktid = oRec("aktid")
	    
	    end if
	    oRec.close
	    
	    dbfunc = "dbred"
	    
	    bdrCol = "red"
	    bgCol = "#ffffe1"
	    
	    oskrift = tsa_txt_252
	    
	
	else
	    
	    aktid = 0
	    matid = 0
	    matantal = 1
	   
	    matnavn = ""
	    matvarenr = 0
	    matkobspris = ""
	    matsalgspris = 0
	    matenhed = "Stk."
	    matgrp = 0
	    forbrugsdato = formatdatetime(date, 2)
	    valuta = basisValId '1 'grundvaluta
	    
	    '*** Til viderefakturering **'
        select case lto 
        case "synergi1", "xintranet - local", "epi", "epi_as", "epi_se", "epi_os", "epi_uk"
        intKode = 1 'Intern
        case else 
	    intKode = 2 'Ekstern = viderefakturering
        end select
	    
	    id = id
	    bilagsnr = bilagsnr
	    dbfunc = "dbopr"
	    
	    bdrCol = "#5582d2"
	    bgCol = "#D6DFf5"
	    
	    oskrift = tsa_txt_193
	    
	    select case lto
	    case "immenso", "epi", "epi_no", "epi_sta", "epi_ab"
	    personlig = 1
	    case else
	    personlig = 0
	    end select
	    
	end if %>

	
	<%'call filterheader(20,20,470,pTxt)
         pTxt = tsa_txt_192
      call filterheader_2013(45,0,560,pTxt)
       %>
	<table cellspacing="2" cellpadding="2" border="0" width=100%>
	<!--<form method="post" action="materialer_indtast.asp?id=<%=id %>&aftid=<%=aftid %>&fromdesk=<%=fromdesk %>&lastid=<%=lastid %>&useFM_jobsog=1">-->
    <form>
        <input type="hidden" name="mid" id="Hidden5" value="<%=usemrn%>">
	<tr><td>
        	<b><%=tsa_txt_078 & " "& tsa_txt_236 %>:</b> 
	
	<%if level <= 2 OR level = 6 then%>
	<br /><input id="vispasogluk" name="vispasogluk" type="checkbox" value="1" <%=vispasoglukCHK %> /> <%=tsa_txt_354%>
    <br /><input id="ignprojekgrp" name="ignprojekgrp" type="checkbox" value="1" <%=ignprojekgrpCHK %> /> <%=tsa_txt_074%> 
    <input id="jobsogtrue" name="jobsogtrue" value="1" type="hidden" />
	<%end if %>
	
	<br /><br /><input id="FM_jobsog" name="FM_jobsog" type="text" value="<%=jobsogval %>" style="border:2px yellowgreen solid; width:470px;" placeholder="Søg job" />


         <br /><br />
       
        <!--<textarea id="matreg_jobid2"></textarea>-->

        <b>Job:</b><br /><select id="matreg_jobid" name="jobid" size="7" style="width:480px;">
	<%
        jobsogVal = ""
        call jobliste(vis, vispasogluk, ignprojekgrp, jobsogVal, id) %>
        </select>
	</td>
	</tr>
        <tr>
	<td>
	<%
    jobid = id 
    status = 1 
    vis = 0    
    call aktlisteOptions(jobid, status, vis, aktid) %>

    <br /><b>Aktivitet:</b><br />

        <select id="matreg_aktid" style="width:480px;">
            <%=straktOptionlist %>
        </select>
	</td>
	</tr>
    
 
    
	
   

      <!--  <tr><td align="right"><input type="submit" value="Søg >>" /></td></tr> -->

   
	</table>
    </form>
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	<br /><br />
	
	
	
	
	<!-- On the fly -->
	<%
	tTop = 35
	tLeft = 0
	tWdth = 550
	
	tId = "mof"
    tVzb = "visible"
    tDsp = ""
	
	'call tableDiv(tTop,tLeft,tWdth)
    call tableDivWid(tTop,tLeft,tWdth,tId, tVzb, tDsp)

	 %>
	
	 <h4><%=oskrift %></h4>

<form>
    

        <input type="hidden" id="matreg_test" value="xx">
        <input type="hidden" id="matreg_lto" value="<%=lto %>" />   
        <input type="hidden" id="matreg_matid" value="0" />
        <input type="hidden" id="matreg_aftid" name="aftid" value="0" />
        <input type="hidden" id="matreg_func" value="<%=dbfunc%>" />
         <input type="hidden" id="matregid" name="matregid" value="<%=matregid%>" />

    	<table width=100% cellspacing=0 cellpadding=0 border=0>
     

        <tr><td colspan="2" id="matreg_indlast_err" style="background-color:#ffffff; padding:4px;"></td></tr>
        <tr><td colspan="2" id="matreg_indlast" style="background-color:#ffffff; padding:4px;"></td></tr>

      

    </table>


	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<!--<form name="opret" action="materialer_indtast.asp?func=<%=dbfunc%>&FM_matid=0&onthefly=1&aftid=<%=aftid%>&FM_sog=<%=sogeKri%>&fromsdsk=<%=fromsdsk%>&matregid=<%=matregid%>" method="post">-->
    
        <!--
	<input type="hidden" name="jobid" id="jobid_0" value="<%=id%>">
        -->

    <%if level <= 2 OR level = 6 then %>
    <tr><td colspan="2"><input type="checkbox" value="1" id="FM_vis_passive" name="FM_vis_passive" <%=chkPassive %> /> Vis passive medarbejdere</td></tr>
    <%end if %>

    <tr>
    <td align=right style="padding:10px 0px 0px 0px;">
    
     <font color=red><b>*</b></font> <%=tsa_txt_237 %>:
    </td> 
    <td style="padding:10px 2px 0px 5px;" id="td_selmedarb">

  <%call selmedarbOptions %>
   </td>
   </tr>
       

	<tr>
		<td align=right style="padding:10px 0px 0px 0px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_204 %>:</td>
		<td style="padding:10px 2px 0px 5px;"><input type="text" name="regdato" id="matreg_regdato_0" value="<%=forbrugsdato%>" size="10"></td>
	</tr>

    <%call matStFelter() %>


            <tr id="dv_mat_otf_sb" style="visibility:<%=dvotf_vzb%>; display:<%=dvotf_dsp%>;"><td colspan="2" align=right><br /><br />
	<input type="button" id="matreg_sb" value="<%=tsa_txt_085 %> >>" />
  <br />
    <br />
    &nbsp;</td></tr>
	
  

	</form>
	</table>
	
    
    <!--table div -->
    </div>
	
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
	<%
	else 'fra lager
	%>
	
	<!--<div>-->
	<%
	
	
	%>


        <!-- Lager -->

	
	
	
	<%'call filterheader(20,20,470,pTxt)
      pTxt = tsa_txt_192
    call filterheader_2013(45,0,720,pTxt)%>
         
	<table cellspacing="2" cellpadding="2" border="0" width=100%>
	<!--<form id="sog" name="sog" action="materialer_indtast.asp?aftid=<%=aftid%>&fromsdsk=<%=fromsdsk%>&useFM_jobsog=1&vis=<%=vis%>" method="post">-->
        <form>
             
        <!-- <input type="hidden" name="mid" id="midids" value="<%=usemrn%>">-->


        <%if level <= 2 OR level = 6 then %>
    <tr><td colspan="4"><input type="checkbox" value="1" id="FM_vis_passive" name="FM_vis_passive" <%=chkPassive %> /> Vis passive medarbejdere</td></tr>
    <%end if %>

	
       <tr><td colspan="4" style="padding-bottom:10px;" id="td_selmedarb"><b><%=tsa_txt_237 %>: </b><br />

        <%call selmedarbOptions %>

        </td>

        
     </tr>

	<tr>
	<td>
        	<b><%=tsa_txt_078 & " "& tsa_txt_236 %>:</b> 
	
	<%if level <= 2 OR level = 6 then%>
	<br /><input id="vispasogluk" name="vispasogluk" type="checkbox" value="1" <%=vispasoglukCHK %> /> <%=tsa_txt_354%>
    <br /><input id="ignprojekgrp" name="ignprojekgrp" type="checkbox" value="1" <%=ignprojekgrpCHK %> /> <%=tsa_txt_074%> 
    <input id="jobsogtrue" name="jobsogtrue" value="1" type="hidden" />
	<%end if %>
	
	<br /><input id="FM_jobsog" name="FM_jobsog" type="text" value="<%=jobsogval %>" style="border:2px yellowgreen solid; width:470px;" placeholder="Søg job" />
    
   <br /><br />
       
        <!--<textarea id="matreg_jobid2"></textarea>-->

        <b>Job:</b><br /><select id="matreg_jobid" name="jobid" size="7" style="width:480px;">
	<%
        jobsogVal = ""
        call jobliste(vis, vispasogluk, ignprojekgrp, jobsogVal, id) %>
        </select>
	</td>
	</tr>
        <tr>
	<td>
	<%
    jobid = id 
    status = 1 
    vis = 0    
    call aktlisteOptions(jobid, status, vis, aktid) %>

    <br /><b>Aktivitet:</b><br />

        <select id="matreg_aktid" style="width:480px;">
            <%=straktOptionlist %>
        </select>
	</td>
	</tr>
        </table>

	<!-- filter header sLut -->
	</td></tr></table>
	
    
   

<br /><br />
<h4>Lager:</h4>
        <table cellspacing="2" cellpadding="2" border="0" width=100%>

  
  <%call mat_lager_sogeKri %>
	
	 
	
	</table>
 <!-- table div--></div>
	
<br /><br /><br /><br /><br /><br />
           
        <input type="hidden" id="matreg_test" value="xx">
        <input type="hidden" id="matreg_lto" value="<%=lto %>" />   
        <input type="hidden" id="matreg_matid" value="0" />
        <input type="hidden" id="matreg_aftid" name="aftid" value="0" />
        <input type="hidden" id="matreg_regdato" name="regdato" value="01-01-2002" />
        <input type="hidden" id="matreg_func" value="<%=dbfunc%>" />
        <input type="hidden" id="matregid" name="matregid" value="0" />
          

    	<table width=100% cellspacing=0 cellpadding=0 border=0>
     

        <tr><td colspan="2" id="matreg_indlast_err" style="background-color:#ffffff; padding:4px;"></td></tr>
        <tr><td colspan="2" id="matreg_indlast" style="background-color:#ffffff; padding:4px;"></td></tr>

      

    </table>
	
    <%select case lto
       case "dencker", "jttek", "intranet - local"
        dv_hgt = 1600
       case else
       dv_hgt = 800
        end select %>    


	<div id="dv_mat_lager_list" style="position:relative; visibility:visible; background-color:#ffffff; display:; padding:10px; height:<%=dv_hgt%>px; overflow-y:scroll; z-index:2000000;">
              <table cellpadding="0" cellspacing="0" border="0" width="100%">

                  <%call matregLagerheader(1) %>
                
                  </table>
            </div>
	
	</form>
	

	
	
	
	
   
    
	
    <!-- 
    <form name="matreg_fralager" id='0' action="materialer_indtast.asp?func=dbopr&id=<%=id%>&fromsdsk=<%=fromsdsk%>&vis=<%=vis %>" method="post">
	
        
		<input type="hidden" name="lto" id="lto" value="<%=lto%>">
		<input type="text" name="FM_indlaest" id="Hidden3" value="<%=strIndlaest%>">
		<input type="hidden" name="aftid" id="aftid" value="<%=aftid%>">
		<input type="hidden" name="jobid" id="jobid_<%=id%>" value="<%=id%>">
		<input type="text" name="skiftlagermsgvist" id="skiftlagermsgvist" value="0">
		<input type="hidden" name="FM_sog" id="Hidden4" value="<%=sogeKri%>">
    
  </form>
            -->

             
	


     <%end if %><!-- vis otf eller lager -->
	
	
	
   








	
	<!-- Senetest indtastede 100 -->
	<%if vis = "otf" then
        s100Top = 95
        else
      s100Top = 190
        end if %>


	<div style="position:absolute; left:780px; width:600px; top:0px;">
	
    <%
    call meStamdata(medid)    
        
    pTxt = meTxt & "<br><span style='font-size:11px;'>" & tsa_txt_216 & "</span>" %>
   
	
	<%'call filterheader(15,0,600,pTxt)
      call filterheader_2013(50,0,645,pTxt)%>
	
	
	<form>
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
	
	<!--<form action="materialer_indtast.asp?id=<%=id %>&aftid=<%=aftid %>&fromsdsk=<%=fromsdsk %>&soglastforty=1&vis=<%=vis%>" method="post">-->
        
	    <tr><td>
	    
	    <%if len(trim(request("sogliste"))) <> 0 AND request("sogliste") <> "0" then 
	    sogliste = request("sogliste")
	    useSog = 1
	    else
	    sogliste = ""
	    useSog = 0
	    end if
	    %>
	    
	    Søg:
            <input id="sogliste" name="sogliste" style="width:200px; border:2px yellowgreen solid;" value="<%=sogliste %>" type="text" placeholder="Bilagsnr, job eller mat. navn" />
            <input id="thisfile" value="matreg" type="hidden" />
            <!-- &nbsp;<input id="Submit1" type="submit" value="<%=tsa_txt_078 %> >> " /> --> (% = wildcard)<br />
            
            <!--
            <input id="sogBilagOrJob" name="sogBilagOrJob" value="0" <%=sogBilagOrJobCHK0 %> type="radio" /> Søg i bilgsnr.<br />
            <input id="sogBilagOrJob" name="sogBilagOrJob" value="1" <%=sogBilagOrJobCHK1 %> type="radio" /> Søg på jobnr.<br />
            -->
           
            <br />
            <input name="showonlypers" id="showonlypers" type="checkbox" <%=showonlypersCHK %> /> <%=tsa_txt_320 %>
            
            <%if level = 1 OR (lto = "dencker" OR lto = "jttek") then %>
            <br />
            <input name="vasallemed" id="vasallemed" type="checkbox" <%=vasallemedCHK %> /> Vis for alle medarbejdere

            <%end if %>
                
	    </td></tr>
	   
	</table>
	 </form>
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	
	
	
	
	
	<%if level <= 2 OR level = 6 then
	cspan = 7
	else
	cspan = 6
	end if %>
	
	
	<% 
	tTop = 242
	tLeft = 0
	tWdth = 782
	
	
	tId = "mhi"
    tVzb = "visible"
    tDsp = ""

    tHgt = 800
	tZindex = 9000


	'call tableDiv(tTop,tLeft,tWdth)
    'call tableDivWid(tTop,tLeft,tWdth,tId, tVzb, tDsp)
     call tableDivAbs(tTop,tLeft,tWdth,tHgt,tId, tVzb, tDsp, tZindex)
	
    showonlypers = 0
    vasallemed = 0

     call senstematReg(usemrn, useSog, sogBilagOrJob, showonlypers, vasallemed, sogliste)

        %>
	
	</div>
     </div>
	
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
	
	
	
	
	
	<div id="lagerosigt" style="position:absolute; left:10px; top:100px; width:1px; height:1px; overflow:auto; visibility:hidden; display:none; width:420px;">
	<b><%=tsa_txt_223%>:</b>  (<%=tsa_txt_224%>)<br />
	<table cellspacing=2 cellpadding=1 border=0 width="100%">
	<tr>
	    <td><b><%=tsa_txt_225%></b></td>
	    <td><b><%=tsa_txt_226%></b></td>
	    <td><b><%=tsa_txt_227%></b></td>
	    <td><b><%=tsa_txt_228%></b></td>
	</tr>
	<%
	strSQL = "SELECT mg.navn AS mgnavn, mg.nummer AS mgnr, "_
	&" m.navn AS mnavn, m.matgrp, m.varenr AS mvnr FROM materiale_grp mg "_
	&" LEFT JOIN materialer m ON (m.matgrp = mg.id) "_
	&" WHERE mg.id <> 0 GROUP BY mg.id ORDER BY m.varenr DESC LIMIT 10" 
	
	'Response.Write strSQL
	'Response.flush
	
	'oRec.open strSQL, oConn, 3
	'while not oRec.EOF 
	%>
	<tr>
	<td><=oRec("mgnavn") %></td>
	<td><=oRec("mgnr") %></td>
	<td class=lille><i><=oRec("mnavn") %></i></td>
	<td class=lille><i><=oRec("mvnr") %></i></td>
	</tr>
	
	
	<%
	
	'oRec.movenext
	'wend
	'oRec.close
	
	 
	 %></table>
	
	</div>
            </div><!-- sidediv -->



            
	
	<%end select%>
    

<%end if
    
     
    %>
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
<!--#include file="../inc/regular/footer_inc.asp"-->
