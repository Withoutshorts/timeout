


<%



level = session("rettigheder")



	
	%>
	
	
    
         <%sub hiddenforms%>
	   
         <input id="FM_lastactiveTD" name="FM_lastactiveTD" value="0" type="hidden" />
         <input id="id" name="id" type="hidden" value="0"/>
        <input id="oprFM_usedatokri" name="FM_usedatokri" type="hidden" value="1"/>
        <input id="oprFM_job" name="FM_job" type="hidden" value="<%=jobid %>"/>
        <input id="oprFM_aftale" name="FM_aftale" type="hidden" value="<%=aftid %>"/>
         <!--<input id="jobelaft" name="jobelaft" type="hidden" value="<%=jobelaft %>"/>-->
         <input id="oprFM_kunde" name="FM_kunde" type="hidden" value="<%=kid %>"/>
         
         <input id="oprFM_start_dag" name="FM_start_dag" type="hidden" value="<%=strDag%>"/>
         <input id="oprFM_start_mrd" name="FM_start_mrd" type="hidden" value="<%=strMrd%>"/>
         <input id="oprFM_start_aar" name="FM_start_aar" type="hidden" value="<%=strAar%>"/>
         
          <input id="oprFM_slut_dag" name="FM_slut_dag" type="hidden" value="<%=strDag_slut%>"/>
         <input id="oprFM_slut_mrd" name="FM_slut_mrd" type="hidden" value="<%=strMrd_slut%>"/>
         <input id="oprFM_slut_aar" name="FM_slut_aar" type="hidden" value="<%=strAar_slut%>"/>
         
         <%end sub %>
	
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	

	
	
	
	
    'func = request("func")
    'thisfile = "erp_opr_faktura"
    'print = request("print")
    err = 0
    

  
   
    
    if cint(jobid) = 0 AND cint(aftid) = 0 then
    err = 88
    end if
    
 
    
    if err <> 0 then
    %>

	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<%
	useleftdiv = "h2"
	errortype = err
	call showError(errortype)
    
    Response.End 
    else
	%>
	
	
	
	
	
	
	<!--<div id=knap1 style="position:relative; width:35px; left:2px; top:45px; visibility:visible; border:1px #8cAAe6 solid; background-color:#ffffff; border-bottom:0px; padding:2px 2px 2px 6px; z-index:2;"><a href="#" onClick="showkontakt();" class=rmenu>Info</a></div>-->
	<!--<div id=knap2 style="position:relative; width:74px; left:39px; top:45px; visibility:visible; border:1px #47AB47 solid; background-color:#ffffff; border-bottom:0px; padding:2px 2px 2px 6px; z-index:2;"><a href="#" onClick="showfakhist();" class=rmenu>Faktura hist.</a></div>-->
	<!--<div id=knap3 style="position:ABSOLUTE; width:53px; left:216px; top:45px; visibility:visible; border:1px #5582d2 solid; background-color:#ffffff; border-bottom:0px; padding:2px 2px 2px 6px; z-index:2;"><a href="#" onClick="showlommeregner();" class=rmenu>Beregn</a></div>-->
	
	
	<!-- KONTAKT IKKE AKTIVE LÆNGERE --->
	<div id=kontakt style="position:relative; height:230px; width:280px; left:0px; top:64px; visibility:hidden; display:none; border:0px #8cAAe6 solid; z-index:1; background-color:#ffffff; overflow:scroll;">
	<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#FFFFff">
	
	<tr><td valign=top width=25 style="padding:5px 0px 0px 5px;"><img src="../ill/ac0009-16.gif" alt="Kontakt" /></td>
	<td valign=top style="padding:10px 0px 0px 0px;">
	
	<%
	
	if cint(jobid) <> 0 then 'job valgt
	
	aftnavn = ""
	
	strSQL = "SELECT j.jobnr, j.id, j.jobnavn, j.jobans1, "_
	&" j.jobans2, k.kkundenavn, k.kkundenr, s.navn AS aftnavn, s.varenr AS aftnr ,kid, "_
	&" budgettimer, ikkebudgettimer, jobtpris, fastpris, "_
	&" m1.mnavn AS m1navn, m1.mnr AS m1mnr, m1.init AS m1init, "_
	&" m2.mnavn AS m2navn, m2.mnr As m2mnr, m2.init AS m2init, jobfaktype, usejoborakt_tp "_
	&" FROM job j "_
	&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft)"_
	&" LEFT JOIN medarbejdere m1 ON (m1.mid = jobans1)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE j.id = "& jobid
	oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    
    strJobnavn = oRec("jobnavn")
    strJobnr = oRec("jobnr")
    
    strKnavn = oRec("kkundenavn")
    strKnr = oRec("kkundenr")
    
    aftnavn = oRec("aftnavn")
    
    kid = oRec("kid")
    
    fakbaretimer = oRec("budgettimer")
    ikkefakbaretimer = oRec("ikkebudgettimer")
    
    budget = oRec("jobtpris")
    fastpris = oRec("fastpris")
    
    
    m1navn = oRec("m1navn")
    m1mnr = oRec("m1mnr")
    m1init = oRec("m1init")
    
    m2navn = oRec("m2navn")
    m2mnr = oRec("m2mnr")
    m2init = oRec("m2init")
    
    aftnr = oRec("aftnr")
    
    jobfaktype = oRec("jobfaktype")
    
    usejoborakt_tp = oRec("usejoborakt_tp")
	
    end if
    oRec.close
	%>
	
	
	<%=strKnavn %> (<%=strKnr %>)
	
	</td></tr><tr><td valign=top style="padding:5px px 0px 5px;"><img src="../ill/ac0063-16.gif" alt="Job" /></td>
   <td valign=top style="padding:5px 0px 5px 0px;"><b><%=strJobnavn %> (<%=strJobnr %>)</b>
   
	
	<%if len(trim(fakbaretimer)) <> 0 then 
	fakbaretimer = fakbaretimer
	else
	fakbaretimer = 0
	end if%>
	
	<%if len(trim(ikkefakbaretimer)) <> 0 then 
	ikkefakbaretimer = ikkefakbaretimer
	else
	ikkefakbaretimer = 0
	end if%>
	
	<br />Timer forkalk.: 
        <b><%=formatnumber(fakbaretimer+ikkefakbaretimer, 2) %></b>
        <br />
	
	Budget: <font class="megetlillesort"><%=basisValISO%></font> <b><%=formatnumber(budget, 2) %></b><br />
	
	<%
	if fastpris = 1 then
	fastprisTxt = "Fastpris"
	    if usejoborakt_tp = 1 then
	    fastprisTxt = fastprisTxt & " (aktiviteter)"
	    end if
	else
	fastprisTxt = "Lbn. Timer"
	end if
	
	 %>
	Jobtype: <b><%=fastprisTxt %></b>
	
	<%
	if len(trim(m1navn)) <> 0 then
	Response.write "<br />Jobansv.: " & m1navn & " ("& m1mnr &")"
	    if len(trim(m1init)) <> 0 then
	    Response.Write " - " & m1init
	    end if
	end if
	
	if len(trim(m2navn)) <> 0 then
	Response.write "<br>Jobejer: " & m2navn & " ("& m2mnr &")"
	    if len(trim(m2init)) <> 0 then
	    Response.Write " - " & m2init
	    end if
	end if 
	 
   
	 %>
	
	</td></tr>
	
	
	            <%
	            if len(trim(aftnavn)) <> 0 then
	            %>
	            <tr><td style="padding:2px 0px 0px 5px; border-top:0px #5582d2 dashed; border-bottom:0px #5582d2 dashed;">
                    <img src="../ill/ac0052-16.gif" alt="Job er tilknyttet denne aftale" />   </td><td style="padding:2px 0px 0px 0px; border-top:0px #5582d2 dashed; border-bottom:0px #5582d2 dashed;">
	            <%=aftnavn %> (<%=aftnr %>) </td></tr> 
	            <%
	            end if
	
	else ' Aftale valgt
	
	aftPris = 0
	aftEnheder = 0
	aftStdato = "1-1-2001"
	aftSldato = "1-1-2001"
	
	strSQL = "SELECT s.navn AS aftnavn, s.aftalenr AS aftnr, kid, kkundenavn, kkundenr, "_
	&" s.pris, s.enheder, s.stdato, s.sldato, s.perafg, s.advitype FROM serviceaft s "_
	&" LEFT JOIN kunder k ON (k.kid = s.kundeid) "_
    &" WHERE s.id = "& aftid
	
	'Response.Write strSQL
	'Response.Flush
	
	oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    
    strKnavn = oRec("kkundenavn")
    strKnr = oRec("kkundenr")
    
    aftnr = oRec("aftnr")
    aftnavn = oRec("aftnavn")
    
    aftPris = oRec("pris")
    aftEnheder = oRec("enheder")
    aftStdato = oRec("stdato")
    aftSlDato = oRec("sldato")
    aftperafg = oRec("perafg")
    aftAdviType = oRec("advitype")
    
    kid = oRec("kid")
    
    end if
    oRec.close
    %>
    
    <b><%=strKnavn %> (<%=strKnr %>)</b>
    </td></tr>
    <tr><td valign=top style="padding:2px 0px 0px 5px; border-top:0px #5582d2 dashed; border-bottom:0px #5582d2 dashed;">
    <img src="../ill/ac0052-16.gif" alt="Aftale" />   </td><td style="padding:2px 0px 0px 0px; border-top:0px #5582d2 dashed; border-bottom:0px #5582d2 dashed;">
	<b><%=aftnavn %> (<%=aftnr %>) </b><br />
	 
	<font class="megetlillesort"><%=basisValISO &" "& formatnumber(aftPris, 2) %><br />
	<%=formatnumber(aftEnheder, 2) %> Enh. / Klip<br />
	<%=formatdatetime(aftStdato, 1) %>
	<%if aftperafg <> 1 then %>
	 - <%=formatdatetime(aftSldato, 1) %>
	<%else %>
	&nbsp;
	<%end if %><br />
	<%if aftAdviType = 2 then %>
	Periode baseret
	<%else %>
	Enhed. / Klip baseret
	<%end if %>
	</td></tr> 
	
	<tr>
	<td valign=top style="padding:2px 0px 0px 5px; border-top:0px #5582d2 dashed; border-bottom:0px #5582d2 dashed;">
    <img src="../ill/ac0063-16.gif" alt="Job der er tilknyttet den valgte aftale" />   </td><td valign=top style="padding:2px 0px 0px 0px; border-top:0px #5582d2 dashed; border-bottom:0px #5582d2 dashed;">
	<%
	strSQLjob = "SELECT jobnavn, jobnr, id FROM job WHERE serviceaft = "& aftid & " ORDER BY jobnavn"
	
	'Response.Write strSQLjob
	'Response.Flush
	jobtilkaft = 0
	oRec.open strSQLjob, oConn, 3
	j = 0
	jobtilkaftSQLkri = " OR (jobid = -1 AND shadowcopy = 0 OR "
	while not oRec.EOF
	%>
	<%=oRec("jobnavn") %> (<%=oRec("jobnr")%>)<br />
	<%
	jobtilkaft = jobtilkaft &","& oRec("id")
	jobtilkaftSQLkri = jobtilkaftSQLkri & " jobid = " & oRec("id") & " AND shadowcopy = 0 OR "
	j = j + 1
	oRec.movenext
	wend
	oRec.close
	
	
	if j = 0 then%>
	(Der er ikke tilknyttet job til denne aftale.)
	<%end if %>
	
	</td>
	</tr>
	
	
    <%end if %>
    
    
	</table>
    </div>
	
	
	
	<%
	
	'Response.end
	
	dontshowDD = 1 %>
	<!--#include file="inc/weekselector_s.asp"-->
	<% 
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag                'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut  'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
    %>
	
	
	
<br />
	
	   <table style="width:280px;">
         <form action="erp_opr_faktura_fs.asp?formsubmitted=3&visfaktura=1&visjobogaftaler=1&visminihistorik=1" method="POST">
             <input type="hidden" id="FM_job_tilknyttet_aftale" name="FM_job_tilknyttet_aftale" value=<%=jobtilkaft%> />
       <tr>
       <td>
       
       <b>Type:</b>&nbsp;
		<%
		'** Faktype ***
		select case intType
		case "0"
		selTypNUL = "SELECTED" 
		selTypET = ""
		selTypTO = ""
		strFaktypeNavn = "Faktura"
		case "1"
		selTypET = "SELECTED"
		selTypTO = ""
		selTypNUL = ""
		strFaktypeNavn = "Kreditnota"
		'case "2"
		'selTypTO = "SELECTED"
		'selTypET = ""
		'selTypNUL = ""
		'strFaktypeNavn = "Rykker"
		case else
		selTypNUL = "SELECTED"
		selTypET = ""
		selTypTO = ""
		strFaktypeNavn = "Faktura"
		end select
		%>
		
        <select id="FM_type" name="FM_type" style="width:100px; font-size:11px;">
                <option value=0 <%=selTypNUL %>>Faktura</option>
                <option value=1 <%=selTypET %>>Kreditnota</option>
             <!--<option value=2 <%=selTypTO%>>Rykker</option>-->
         </select> 
       
     
       
       </td>
       
       <td align=center valign=middle style="padding:3px; border:1px #6CAE1C solid; background-color:#DCF5BD;">
       <input id="Submit2" type="submit" value=" Opret ny >> " />
           </td></tr>
	    </table>
	
	
	<!-- Fakturaer mini historik --->
<br /><br />
 <a href="#" id="showfakind" style="color:#999999; font-size:9px; font-weight:lighter;">[ Fak. Pre-indstill. ]</a>

	<div id=fakturaer style="position:relative; height:260px; top:0px; background-color:#ffffff; visibility:visible; display:; width:275px; overflow:scroll; z-index:1;">
	
         
        <table cellpadding=0 cellspacing=0 border=0 width=100%><%
	 
	    if cint(igDato) = 1 then
        SQLkriPer = ""
	    else
	    SQLkriPer = " AND ((f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') OR (f.labeldato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' AND brugfakdatolabel = 1)) "
	    end if
	    
	    
	    lastFak = "2001/01/01"
		fakfindes = 0
		lastFaknr = 0
		
		
		if jobid = 0 then
		jobaftSQL = "f.aftaleid = "& aftid &""
		len_jobtilkaftSQLkri = len(jobtilkaftSQLkri)
		left_jobtilkaftSQLkri = left(jobtilkaftSQLkri, len_jobtilkaftSQLkri - 3)
		jobtilkaftSQLkri = left_jobtilkaftSQLkri & ")"
		
		jobaftSQL = jobaftSQL & jobtilkaftSQLkri
		
		else
		jobaftSQL = "f.jobid = "& jobid &""
		end if
		
		strSQLFak = "SELECT f.jobid, f.aftaleid, f.fid, f.fakdato, f.faknr, f.betalt, "_
		&" f.faktype, f.beloeb, f.shadowcopy, f.valuta, f.kurs, v.valutakode, brugfakdatolabel, labeldato, medregnikkeioms, fak_laast "_
		&" FROM fakturaer f LEFT JOIN valutaer v ON (v.id = f.valuta) "_
		&" WHERE ("& jobaftSQL &")"& SQLkriPer &""_
		&" ORDER BY f.faknr DESC"
	        
	        'Response.Write strSQLFak &"<br>" 
	        'Response.Flush
	        
		f = 1
		lastFakfundet = 0
		oRec3.open strSQLFak, oConn, 3
        while not oRec3.EOF 
            
            if f = 1 then 
            fakfindes = 1
            end if
            
            select case oRec3("faktype")
            case 0
                if cint(lastFakfundet) <> 1 AND cint(oRec3("medregnikkeioms")) <> 1 AND cint(oRec3("medregnikkeioms")) <> 2 then
                lastFakfundet = 1
                lastFak = oRec3("fakdato")
                lastFaknr = oRec3("faknr")
                end if
            tyThis = "Fak."
            belthis = oRec3("beloeb")
            case 1
            tyThis = "Kre."
            belthis = -(oRec3("beloeb"))
            'case 2
            'tyThis = "Ryk."
            'belthis = oRec3("beloeb")
            end select
            
            if cint(id) = oRec3("fid") then
            bgthis = "#ffff99"
            else
            bgthis = "#ffffff"
            end if
            
            if oRec3("betalt") = 1 then
            acls = "erp_green"
            else
            acls = "erp_silver"
            end if
            
            %>
            <tr>
            <%if (oRec3("shadowcopy") <> 1 AND jobid <> 0) OR (aftid <> 0 AND oRec3("aftaleid") <> 0) then %>
            
            <td id="td_1_<%=f%>" style="background-color:#ffffff; border-bottom:1px #cccccc dashed; padding:0px 2px 0px 2px; white-space:nowrap;" class=lille style="background-color:#ffffff;">
            
            <%if oRec3("brugfakdatolabel") = 1 then%>
            L: <%=replace(formatdatetime(oRec3("labeldato"),2),"-",".") %>
            <%else %>
            F: <%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %>
            <%end if %>
            &nbsp;<%=tyThis%>


            <%if oRec3("medregnikkeioms") = 1 then %>
            (i)
            <%end if %>

            <%if oRec3("medregnikkeioms") = 2 then %>
            (h)
            <%end if %>
            
            </td>
            
            <td id="td_3_<%=f%>" class=lille style="border-bottom:1px #cccccc dashed; white-space:nowrap;"><a href="erp_opr_faktura_fs.asp?func=red&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=2&visjobogaftaler=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>" class=<%=acls %> onclick="activetd(<%=f%>)">&nbsp;<%=oRec3("faknr")%></a></td>
            <td id="td_4_<%=f%>" style="background-color:#ffffff; border-bottom:1px #cccccc dashed; height:20px; white-space:nowrap; padding:2px 2px 0px 2px;">

                <!-- erp_fak_godkendt_2007.asp?id=<%=oRec3("fid")%>&FM_usedatointerval=1&jobid=<%=jobid%>&aftid=<%=aftid%>&kid=<%=kid%>&FM_start_dag_ival=<%=strDag%>&FM_start_mrd_ival=<%=strMrd %>&FM_start_aar_ival=<%=strAar%>&FM_slut_dag_ival=<%=strDag_slut%>&FM_slut_mrd_ival=<%=strMrd_slut%>&FM_slut_aar_ival=<%=strAar_slut%> -->
            <%
            if oRec3("betalt") <> 1 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) then%>
            <a href="erp_opr_faktura_fs.asp?func=red&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=1&visjobogaftaler=1&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>"><img src="../ill/redigerfak.gif" border="0" alt="Rediger faktura" /></a>
                <!--  -- >
            <!-- &rykkreopr=j -->
            <%
                if oRec3("fak_laast") <> 1 then ' har faktura været låst?? 
                %><a href="erp_opr_faktura_fs.asp?func=slet&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=0&visjobogaftaler=1"><img src="../ill/sletfak.gif" border="0" alt="Slet faktura" /></a><br />
                <%end if %>

                <!-- &FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%> -->

            <%else %>
               
               
                <%
                select case lto
                case "xsynergi1"

                  if session("mid") = 1 then%>
                  
                
                <a href="erp_opr_faktura_fs.asp?func=fortryd&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=0&visjobogaftaler=1" class=vmenualt>
                    <img src="../ill/fortryd.gif" border="0" alt="Fortryd godkend" /></a>
                   
                
                <%end if 

                case else
                
                if level = 1 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) then%>
                  
                <!-- Skal IKKE være muligt pga huller i faknr reække følge ELLER HVAD:: ***' --> 
                <a href="erp_opr_faktura_fs.asp?func=fortryd&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=0&visjobogaftaler=1" class=vmenualt>
                    <img src="../ill/fortryd.gif" border="0" alt="Fortryd godkend" /></a>
                   
                
                <%end if 
                
                end select%>

                 &nbsp;
                
              
                
            <%end if %>
            
                </td>
                <td align=right id="td_2_<%=f%>" style="background-color:#ffffff; border-bottom:1px #cccccc dashed; padding:0px 2px 0px 2px; white-space:nowrap;" class=lille><%=formatnumber(belthis, 0) &" "& oRec3("valutakode") %></td>
            
            
            <%
            
            call beregnValuta(belthis,oRec3("kurs"),100)
            totbel = totbel + (valBelobBeregnet)
            
            else 'shadowcopy %>
                
                <td id="td_1_<%=f%>" style="background-color:#ffffff; border-bottom:1px #cccccc dashed; padding:0px 2px 0px 2px; white-space:nowrap;" class=lillegray style="background-color:#ffffff;">
                 <%if oRec3("brugfakdatolabel") = 1 then%>
            L: <%=replace(formatdatetime(oRec3("labeldato"),2),"-",".") %>
            <%else %>
            F: <%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %>
            <%end if %>
                
              &nbsp;<%=tyThis%>
              
               <%if oRec3("medregnikkeioms") = 1 then %>
            (i)
            <%end if %>
              </td>
                <td id="td_3_<%=f%>" class=lillegray style="border-bottom:1px #cccccc dashed; white-space:nowrap;">&nbsp;<%=oRec3("faknr")%></td>
                <td colspan=2 id="td_4_<%=f%>" class=lillegray  style="background-color:#ffffff; border-bottom:1px #cccccc dashed;">
                
                
                <% 
                
                '*** Hvis man er inde på et job og skal de de
                '*** fakturaer der ligger på aftaler som dettet job er tilknyttet **
                if jobid <> 0 then
                
                aftnavn = ""
                aftnr = ""
                
                strSQLsc = "SELECT f.aftaleid, s.navn, s.aftalenr FROM fakturaer f "_
                &"LEFT JOIN serviceaft s ON (s.id = f.aftaleid) WHERE f.faknr = "& oRec3("faknr") &" AND f.shadowcopy = 0"
                
                oRec.open strSQLsc, oConn, 3
                if not oRec.EOF then
                
                aftnavn = oRec("navn")
                aftnr = oRec("aftalenr")
                
                end if
                oRec.close%>
                
                    <%if len(aftnavn) > 12 then %>
                    <%=left(aftnavn, 12) & ".. ("& aftnr &")" %>
                    <%else %>
                    <%=aftnavn & " ("& aftnr &")" %>
                    <%end if %>
                
                <%else 
                
                jobnavn = ""
                jobnr = ""
                
                strSQLsc = "SELECT f.jobid, j.jobnavn, j.jobnr FROM fakturaer f "_
                &"LEFT JOIN job j ON (j.id = f.jobid) WHERE f.faknr = "& oRec3("faknr") &" AND f.shadowcopy = 0"
                
                oRec.open strSQLsc, oConn, 3
                if not oRec.EOF then
                
                jobnavn = oRec("jobnavn")
                jobnr = oRec("jobnr")
                
                end if
                oRec.close
                %>
                
                    <%if len(jobnavn) > 12 then %>
                    <%=left(jobnavn, 12) & ".. ("& jobnr &")" %>
                    <%else %>
                    <%=jobnavn& " ("& jobnr &")" %>
                    <%end if %>
                    
                
                
            <%end if %>
                </td>
                
                
            <%
            
            totbel = totbel 'shadowcopy
            end if %>
            </tr>
            <% 
            
       
        f = f + 1
        oRec3.movenext
        wend 
        oRec3.close
                
        
       
       %>
        <tr>
            <td colspan=4 align=right style="padding:10px 8px 0px 2px;">
            <%if f = 1 then %>
            (Der findes ingen fakturaer i det valgte interval.)
            <%else %>
            Ialt: <b> <%=formatnumber(totbel, 2) &" "& basisValISO%></b> <br />Inkl. kladder og interne. Ekskl. moms</td></tr>
            <%end if %>
        
        <tr><td colspan=4 style="padding:10px 2px 0px 5px;">
        
          <%
                
                if cdate(strDag &"/"& strMrd &"/"& strAar) > cdate(lastFak) then
                lastFakuseSQL = strAar &"/"& strMrd &"/"& strDag
                else
                lastFakuseSQL = dateadd("d", 1, lastFak)
                lastFakuseSQL = year(lastFakuseSQL) &"/"& month(lastFakuseSQL) &"/"& day(lastFakuseSQL)
                end if
          
          
          
          if jobid <> 0 then
          
          
               
                
                sumTimerVedFak = 0
                
                if fakfindes = 1 then
                strSQLtimer = "SELECT sum(timer) AS sumtimerEfterFak FROM timer WHERE "_
                &" tfaktim <> 5 AND tjobnr = '"& strJobnr &"'"_ 
	            &" AND tdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY tjobnr"
	            
	            'Response.Write strSQLtimer &"<br>"
	            'Response.flush
	            
	            oRec2.open strSQLtimer, oConn, 3
                if not oRec2.EOF then
                
                sumTimerVedFak = oRec2("sumtimerEfterFak")
                
                end if
                oRec2.close
                
                sumMattilFak = 0
                strSQLmat = "SELECT sum(matantal) AS matantal FROM materiale_forbrug WHERE jobid = "& jobid &" AND "_
                &" forbrugsdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY matid"
	            oRec2.open strSQLmat, oConn, 3
                if not oRec2.EOF then
                
                sumMattilFak = oRec2("matantal")
                
                end if
                oRec2.close 
	                      
                
                Response.Write "<br><b>Til fakturering</b> (siden sidste fak., ikke interne)<br>Timer:  <b> " & sumTimerVedFak & "</b><br>"
	            Response.Write "Materialer: <b>"& sumMattilFak &"</b><br><br>&nbsp;"
     
      
                end if
     
     
            else 'Aftale 
      
                 '** Enheder 
                                    
                sumEnhVedFak = 0
                strSQLenh = "SELECT sum(timer * faktor) AS sumEnhEfterFak FROM timer "_
                &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE "_
                &" tfaktim <> 5 AND seraft = "& aftid &""_ 
                &" AND tdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY seraft, taktivitetid"
	            
                'Response.Write strSQLenh &"<br>"
	            
                oRec3.open strSQLenh, oConn, 3
                while not oRec3.EOF
                
                sumEnhVedFak = sumEnhVedFak + oRec3("sumEnhEfterFak")
                
                oRec3.movenext
                wend
                oRec3.close
                
                 Response.Write "<br><b>Til fakturering</b> (siden sidste fak. ikke interne) <br>Enheder: <b> " & sumEnhVedFak & "</b><br><br>&nbsp;"

            end if%>
      </td></tr>
      
      <%call hiddenforms %>
      </form>
    </table>

         

               

              
              
        </div>


	
	
       
        
        
        
	
	
	 <div id="fakind_d" style="position:absolute; width:275px; height:500px; left:-20px; top:114px; border:10px #CCCCCC solid; visibility:hidden; display:none; z-index:10000; padding:10px 10px 10px 10px; background-color:#ffffff;">
        <table cellpadding=0 cellspacing=4 border=0 width=100%>
        <form id=prefak action="erp_opr_faktura_fs.asp?formsubmitted=3&visfaktura=1&visjobogaftaler=1&visminihistorik=1&func=opdprefak" method="post">
            <tr><td align="right" style="padding:10px 10px 10px 10px;"><span id="luk_fakpre" style="color:red;">[X]</span></td></tr>
        <tr><td style="padding:10px 10px 10px 10px;">
        <h4>Faktura pre. indstillinger:</h4>
        
        <b>A)</b> Visning af medarbejder timer
		på faktura print skal automatisk være slået til:<br />
		<%
		if request.cookies("erp")("tvmedarb") = "1" then
		chkA = "CHECKED"
		chkB = ""
		else
		chkB = "CHECKED"
		chkA = ""
		end if
		%>
		<input type="radio" name="FM_chkmed" value="1" <%=chkA%>> Ja 
		<input type="radio" name="FM_chkmed" value="0" <%=chkB%>> Nej
		
		<br /><br />
		<b>B)</b> Forvælg "vis joblog og materiale log" på faktura print? <br />
		<%
		if request.cookies("erp")("tvlogs") = "1" then
		chklogA = "CHECKED"
		chklogB = ""
		else
		chklogB = "CHECKED"
		chklogA = ""
		end if
		%>
		<input type="radio" name="FM_chklog" value="1" <%=chklogA%>> Ja 
		<input type="radio" name="FM_chklog" value="0" <%=chklogB%>> Nej
		
		
		<br /><br />
		<b>C)</b> Forvælg "vis tlf." på faktura print? <br />
		<%
		if request.cookies("erp")("tvtlffax") = "1" then
		chktlffaxA = "CHECKED"
		chktlffaxB = ""
		else
		chktlffaxB = "CHECKED"
		chktlffaxA = ""
		end if
		%>
		<input type="radio" name="FM_chktlffax" value="1" <%=chktlffaxA%>> Ja 
		<input type="radio" name="FM_chktlffax" value="0" <%=chktlffaxB%>> Nej
		
		<br /><br />
		<b>D)</b> Forvælg "vis email adresse" faktura print? <br />
		<%
		if request.cookies("erp")("tvemail") = "1" then
		chkemailA = "CHECKED"
		chkemailB = ""
		else
		chkemailB = "CHECKED"
		chkemailA = ""
		end if
		%>
		<input type="radio" name="FM_chkemail" value="1" <%=chkemailA%>> Ja 
		<input type="radio" name="FM_chkemail" value="0" <%=chkemailB%>> Nej
		
		<!--
		<br /><br />
		<b>E)</b> Vis alle medarbejdere uanset om de er de-aktiverede eller tilknyttet via projektgrupper. <br />
		Der kan forekomme lang load tid.<br />
		<%
		'if request.cookies("erp")("deak") = "1" then
		'chkdeakA = "CHECKED"
		'chkdeakB = ""
		'else
		'chkdeakB = "CHECKED"
		'chkdeakA = ""
		'end if
		%>
		<input type="radio" name="FM_chkdeak" value="1" <%=chkdeakA%>> Ja 
		<input type="radio" name="FM_chkdeak" value="0" <%=chkdeakB%>> Nej
		-->
		
		
		
		
		<br /><br />
		<b>E)</b> Slå "luk job ved fakturering" til.<br />
		<%
		if request.cookies("erp")("lukjob") = "1" then
		chklukjobA = "CHECKED"
		chklukjobB = ""
		else
		chklukjobB = "CHECKED"
		chklukjobA = ""
		end if
		%>
		<input type="radio" name="FM_chklukjob" value="1" <%=chklukjobA%>> Ja 
		<input type="radio" name="FM_chklukjob" value="0" <%=chklukjobB%>> Nej


        <br /><br />
        <%if cint(ign_akttype_inst) = 1 then
        ign_akttype_inst_CHK = "CHECKED"
        else
        ign_akttype_inst_CHK = ""
        end if %>
        <b>F)</b>  <input id="Checkbox1" name="ign_akttype_inst" value="1" type="checkbox" <%=ign_akttype_inst_CHK %> /> Ignorer akttype-instillinger (vis alle aktiviteter uanset om typen er sat til at blive vist på faktura)
	  

		
		 <%call hiddenforms %>
	    
	    <br />
            &nbsp;
	    </td>
	    
	    </tr>
         

            <tr><td align=right style="padding:10px 10px 10px 10px;">
            <input id="Submit3" type="submit" value="Opdater >>" />
	    </td></tr>
		</form>
		</table>

     

		</div>
	
	
	    
	   
       
        
        
      
     




     
     <!--- Lommeregner --->
	<div id=lommeregner style="position:absolute; background-color:#ffffff; width:275px; height:230px; top:64px; border:0px #5582d2 solid; visibility:hidden; display:none; z-index:1;">
	<table cellspacing=0 cellpadding=0 border=0><form id=beregn name=beregn>
	<tr>
		<td style="padding:10px 2px 0px 10px;" valign=top>Beløb: <input type="text" name="beregn_belob" id="beregn_belob" value="0" size="4" style="font-size:9px;"> <b>/</b> </td>
	    <td style="padding:10px 2px 0px 2px;" valign=top>Timer: <input type="text" name="beregn_timer" id="beregn_timer" value="0" size="4" style="font-size:9px;"></td>
	    <td style="padding:10px 2px 0px 2px;" valign=top><input type="button" name="beregn" id="beregn" value=" = " onClick="beregntimepris()" style="font-size:9px;">&nbsp; <input type="text" name="beregn_tp" id="beregntp" value="0" style="width:58px; font-size:9px;"></td>
	</tr></form></table>								
	</div>
	<!-- ASLUT lommeregner -->
	    
    
    
      <!--
     <br />
	 <a href="erp_fak.asp?FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&jobelaft=<%=jobelaft%>&FM_kunde=<%=kid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" class=vmenu target="_blank">Opret Faktura på valgte job/aftale og med den angivne fkatura dato.</a>
	 -->
 
   
	<%end if 'validering%>
	