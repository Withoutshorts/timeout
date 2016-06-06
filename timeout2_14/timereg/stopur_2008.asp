	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<% 
	'*********** Sætter sidens global variable ***********************************
	
	'* Hvis der vælges en anden bruger
	'* den den der er logget på.
	
	if len(request("reload")) <> 0 then
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	end if
	 
	 '*** Mid !***
	if len(Request("FM_mid")) <> 0 then
	medid = Request("FM_mid")
	else
	    if request.Cookies("stopur")("medid") <> "" then
	    medid = request.Cookies("stopur")("medid")
	    else
	    medid = session("Mid")
	    end if
	end if
	
	Response.Cookies("stopur")("medid") = medid
	
	thisfile = "stopur_2008.asp"
	func = request("func")
	
	level = session("rettigheder")
	
	if len(request("incid")) <> 0 then
	incid = request("incid")
	else
	incid = 0
	end if
	
	if len(request("logentry")) <> 0 then
	logentry = request("logentry")
	else
	logentry = 0
	end if
	
	if len(request("print")) <> 0 then
	print = request("print")
	else
	print = ""
	end if
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid") 
	else
	    if request.Cookies("stopur")("jobid") <> "" AND (len(request("sorter")) <> 0 OR print = "j") then
	    jobid = request.Cookies("stopur")("jobid")
	    else
	    jobid = 0
	    end if
	end if
	
	Response.Cookies("stopur")("jobid") = jobid
	
	if len(request("aktid")) <> 0 then
	aktid = request("aktid")
	else
	aktid = 0
	end if
	
	if len(request("FM_besk")) <> 0 then
	kommentar = replace(request("FM_besk"), "'", "''")
	else
	kommentar = ""
	end if
	

	strEditor = session("user")
	strDato = session("dato")
	datoSQL = year(strDato) &"/"& month(strDato) &"/"& day(strDato)

	
		if len(request("sorter")) <> 0 then
		sorter = request("sorter")
		else
			if request.cookies("stopur")("sorter") <> "" then
			sorter = request.cookies("stopur")("sorter")
			else
			sorter = 0
			end if
			
		end if 
		
		Response.Cookies("stopur")("sorter") = sorter
			

		select case sorter
		case 1
		sCHK0 = ""
		sCHK1 = "CHECKED"
		sCHK2 = ""
		orderBySQL = " s.id"
		case 2
		sCHK0 = ""
		sCHK1 = ""
		sCHK2 = "CHECKED"
		orderBySQL = " s.medid, s.sttid"
		case else
		sCHK0 = "CHECKED"
		sCHK1 = ""
		sCHK2 = ""
		orderBySQL = " s.sttid"
		end select
	
	
	
	if len(request("sorter")) <> 0 then
	    
	    if len(request("vis")) <> 0 then 
	    vCHK = "CHECKED"
	    vis = request("vis")
	    visSQL = ""
	    else
	    vis = 0
		vCHK = ""
		visSQL = " AND timereg_overfort <> 1 "
	    end if
	    
	else
		
		if request.cookies("stopur")("vis") <> "" then
		
		vis = request.cookies("stopur")("vis")
		
			if vis = 1 then
			visSQL = ""
			vCHK = "CHECKED"
			else
			visSQL = " AND timereg_overfort <> 1 "
			vCHK = ""
			end if
			
		else
		vis = 0
		vCHK = ""
		visSQL = " AND timereg_overfort <> 1 "
		end if
		
	end if
	
	if len(request("sorter")) <> 0 then
	
	    if len(request("FM_brugPer")) <> 0 then
	    brugPer = 1
	    brugPerCHK = "CHECKED"
	    else
	    brugPer = 0
	    brugPerCHK = ""
	    end if
    	
	   
	else
	    
	    brugPer = Request.Cookies("stopur")("brugPer")
	    
	    if brugPer = "1" then
	    brugPerCHK = "CHECKED"
	    else
	    brugPerCHK = ""
	    end if
	     
	
	end if
	
	Response.Cookies("stopur")("brugPer") = brugPer
	
	
	
	
	Response.Cookies("stopur")("vis") = vis
	Response.Cookies("stopur").expires = date + 10	

	%>
	<script>
	
	function checkAll()
	{
	antal_t = document.getElementById("antalsids_t").value
	for (i = 0; i < antal_t; i++)
		document.getElementById("tilvalgt_"+i+"").checked = true ;
	}
	
	function uncheckAll()
	{
	antal_t = document.getElementById("antalsids_t").value
	for (i = 0; i < antal_t; i++)
		document.getElementById("tilvalgt_"+i+"").checked = false ;
	}
	
	function subm(url){
	alert(url)
	window.location.href = url
    }
	
	</script>
	<%
        	
	    function SQLBless(s)
		        dim tmp
		        tmp = s
		        tmp = replace(tmp, ",", ".")
		        SQLBless = tmp
        end function

        function SQLBless2(s)
		        dim tmp2
		        tmp2 = s
		        tmp2 = replace(tmp2, ".", ",")
		        SQLBless2 = tmp2
        end function
        	
	
	
	select case func
	case "slet"
	    
	    sid = request("sid") 
	    strSQLdel = "DELETE FROM stopur WHERE id = " & sid
	    oConn.execute(strSQLdel)
	    
	    Response.Redirect "stopur_2008.asp" 
	    
	case "fortryd"
	    
	    sid = request("sid") 
	    strSQLdel = "UPDATE stopur SET sid_godkendt = 0 WHERE id = " & sid
	    oConn.execute(strSQLdel)
	    
	    Response.Redirect "stopur_2008.asp?incid="&incid&"&aktid="&aktid
	    
	case "ins"
	    
	    sttidSQL = year(now) &"/"& month(now) &"/"& day(now) &" "& formatdatetime(now, 3)
	    
	    
	    strSQLins = "INSERT INTO stopur (sttid, incident, incidentlog, medid, jobid, aktid, dato, editor, kommentar) VALUES"_
	    &" ('"& sttidSQL &"', "& incid &", "& logentry &", "& medid &", "& jobid &", "& aktid &", '"& datoSQL &"', '"& strEditor &"', '"& kommentar &"')"
	    
	    'Response.Write strSQLins
	    'Response.end
	    
	    oConn.execute(strSQLins)
	    
	    Response.Redirect "stopur_2008.asp?incid="&incid&"&aktid="&aktid
	     
	case "opd"
	
	
	 '*** Sletter log entrys ****'
	    'strSQLdel = "DELETE FROM stopur WHERE incidentlog = " & logentry
	    'oConn.execute(strSQLdel)
	    
	    '** Opretter log entrys ****'
	    sids = split(request("antalsids"), ", ")
	    'Response.Write request("antalsids")
	    'Response.flush
	    
	    for t = 0 to UBOUND(sids)
	    
	    'Response.Write sids(t) & "<br>"
	    
	    if len(trim(request("FM_sttid_"&sids(t)&""))) <> 0 AND isdate(request("FM_sttid_"&sids(t)&"")) then
	    sttid = request("FM_sttid_"&sids(t)&"")
	    sttidSQL = year(sttid) &"/"& month(sttid) &"/"& day(sttid) &" "& formatdatetime(sttid, 3)
	        
	        if request("hentgem_"&sids(t)&"") = 0 then
	        sltidSQL = year(now) &"/"& month(now) &"/"& day(now) &" "& formatdatetime(now, 3)
	        else
	        sltid = request("FM_sltid_"&sids(t)&"")
	        
	            
	            if isdate(request("FM_sltid_"&sids(t)&"")) then
	            sltidSQL = year(sltid) &"/"& month(sltid) &"/"& day(sltid) &" "& formatdatetime(sltid, 3)
	            else
	            sltidSQL = ""
	            end if
	            
	        
	        end if
	        
	    else
	    sttidSQL = ""
	    sltidSQL = ""
	    end if
	    
	    if len(request("sid_godkendt_"&sids(t)&"")) <> 0 then
	    sid_godkendt = request("sid_godkendt_"&sids(t)&"")
	    else
	    sid_godkendt = 0
	    end if
	    
	    if sttidSQL <> "" AND sltidSQL <> ""  then
	    strSQLins = "UPDATE stopur SET sttid = '"& sttidSQL &"', sltid = '"& sltidSQL &"', aktid = "& aktid &", "_
	    &" sid_godkendt = "& sid_godkendt &", dato = '"& datoSQL &"', editor = '"& strEditor &"', kommentar = '"& kommentar &"' WHERE id = " & sids(t)
	    
	    '&" medid = "& session("mid") &", incident = "& incid &", incidentlog = "& logentry &""_
	    
	    'Response.Write strSQlins &"<br>"
	    'Response.flush
	    oConn.execute(strSQLins)
	    else
	    
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	    <%
	    errortype = 116
	    call showError(errortype)
	    Response.end
	    
	    end if
	    
	    next
	    
	    
	    Response.Redirect "stopur_2008.asp?incid="&incid&"&aktid="&aktid
	
	case else
	
	
    '*** Henter projektgrupper ***'
    call hentbgrppamedarb(medid) 
	    
	 
	    
	  



'*******************************************************************************
%>
<script>
function kopiertxt(t){
txt = document.getElementById("FM_hidden_txt_"+t+"").value
document.getElementById("FM_besk_"+t+"").value = txt
}

</script>

 <body onLoad="self.focus()">
	<div id="sindhold" style="position:relative; padding:0px; width:920px; left:10px; top:20px; visibility:visible;">
	
	<h4><img src="../ill/stopur_48.png" /> TimeOut Stopur Entries</h4>
	
	<%call filterheader(0,0,600,pTxt)%>
	<table cellpadding=2 cellspacing=0 border=0 width=100% bgcolor="#ffffff">
	<form method=post action="stopur_2008.asp">
	<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td colspan=2><b>Periode:</b>
			<%if print = "j" then
			dontshowDD = 1
			Response.Write "&nbsp;"
			else
			Response.Write "<br>"
			end if %>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			<%if print = "j" then%>
			<%=replace(formatdatetime(strDag&"/"&strMrd&"/"&strAar, 2), "-", ".") %> til  <%=replace(formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2), "-", ".")%>
			<%end if %>
			
			<%if print <> "j" then%>
                <br /><input id="Checkbox2" name="FM_brugPer" value="1" type="checkbox" <%=brugPerCHK%> /> Vis kun entries i valgte periode (ifht. starttid på entry):
			<%else %>
			    
			    <br /> Vis kun entries i valgte periode (ifht. starttid på entry): 
			    
			    <%if brugPer = 1 then %>
			    &nbsp;<b>Ja</b>
			    <%else %>
			    &nbsp;<b>Nej</b>
			    <%end if %>
			<%end if %>
			</td>
	</tr>
	</td>
	
	</tr>
	<tr>
	    <td><br />
	        <b>Medarbejder:</b>&nbsp;
	        
	        <%if level < 2 then %>
	        <%if print <> "j" then %>
	        <select id="FM_mid" name="FM_mid" style="width:200px; font-size:10px;">
	        <option value=0>Alle</option>
	        <%else 
	        strMedPrint = "Alle"
	        end if 
	        
	       
	        
	        
	        strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF 
	        
	        if cint(medid) = oRec("mid") then
	        selMid = "SELECTED"
	        strMedPrint = oRec("mnavn") &" ("& oRec("mnr") &")"
	        else
	        selMid = ""
	        end if
	        
	        if print <> "j" then %>
	        <option value="<%=oRec("mid") %>" <%=selMid %>><%=oRec("mnavn")%> (<%=oRec("mnr") %>)</option>
	        <%end if
	        
	        oRec.movenext
	        wend
	        oRec.close
	        if print <> "j" then %>
	        </select>
	        <%else %>
	        <%=strMedPrint %>
	        <%end if %>
	        <%else %>
            <input id="FM_mid" name="FM_mid" value="<%=medid %>" type="hidden" />
	        <%end if %>
	        <br /><br />
            
            <b>Kontakt / job:&nbsp;</b>
	        
	        <%if level < 2 then %>
	        <%if print <> "j" then %>
	        <select id="Select2" name="jobid" style="width:400px; font-size:10px;">
	        <option value=0>Alle</option>
	        <%else 
	        strJobPrint = "Alle"
	        end if
	        
	        strSQL = "SELECT k.kkundenavn, k.kkundenr, j.jobnavn, j.jobnr, j.id AS jid FROM job j "_
	        &" LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE j.id <> 0 AND j.jobstatus = 1 ORDER BY k.kkundenavn, j.jobnavn"
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF 
	        
	        if cint(jobid) = oRec("jid") then
	        selJid = "SELECTED"
	        strJobPrint = oRec("kkundenavn")&" ("&oRec("kkundenr")&") | "& oRec("jobnavn") &" ("& oRec("jobnr")&")"
	        else
	        selJid = ""
	        end if
	        
	        if print <> "j" then %>
	        <option value="<%=oRec("jid") %>" <%=selJid %>><%=oRec("kkundenavn")%> (<%=oRec("kkundenr") %>) | <%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)</option>
	        <%end if
	        
	        oRec.movenext
	        wend
	        oRec.close
	        if print <> "j" then%>
	        </select> (kun aktive job)
	        <%else %>
	        <%=strJobPrint %>
	        <%end if %>
	        
	        <%else %>
            <input id="Hidden1" name="FM_jobid" value="<%=jobid%>" type="hidden" />
	        <%end if %>
	        <br /><br />
		    
		    <%if print <> "j" then %>
            <input id="sorter" name="sorter" value=0 type="radio" <%=sCHK0 %>/>Sorter efter dato.<br />
	       <input id="sorter" name="sorter" value=1 type="radio" <%=sCHK1 %>/>Sorter efter Incident Id.<br />
	       <input id="sorter" name="sorter" value=2 type="radio" <%=sCHK2 %>/>Sorter efter Medarbejder<br />
		    <input id="vis" type="checkbox" name="vis" value="1" <%=vCHK %> /> Vis overførte entries.
		    <%end if%>
	   </td>
	</tr>
	<%if print <> "j" then %>
	<tr>
	    <td align=right>
            <input id="Submit1" type="submit" value="Vis periode >> " /></td>
	</tr>
	<%end if%>
	</form>
	</table>
	
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	
	
	<div id="sindhold2" style="position:relative; left:0px; top:20px; visibility:visible;">
	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 940
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	<table cellpadding=2 cellspacing=0 border=0 width=100% bgcolor="#ffffff">
	<!--<tr>
		
		<td colspan=13 bgcolor="#d6dff5">
		<a href="#" name="CheckAll" onClick="checkAll()"><img src="../ill/alle.gif" border="0"></a>
		&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll()">
		<img src="../ill/ingen.gif" border="0"></a>
		</td>
		
	</tr>-->
	<tr bgcolor="#5582d2">
	    <td class=alt><b>Id</b></td>
	    <td class=alt width=75><b>Status</b></td>
	    <td class=alt><b>Starttid</b></td>
	    <td class=alt><b>Sluttid</b></td>
	    <td class=alt width=50 align=right>Timer:Min.</td>
	    <%if print <> "j" then %>
	    <!--<td class=alt>Hent ny slut-tid</td>
	    <td class=alt>Gem ind-tast.</td>-->
	    <%else %>
	    <!--<td class=alt>
            &nbsp;</td>
        <td class=alt>
            &nbsp;</td>-->
            
	    <%end if %>
	    
	    <td class=alt style="padding-left:10px;">Kunde<br><b>Job</b><br />
	    Aktivitet</td>
	   
	    <td class=alt><b>Incident Id</b></td>
	    <td class=alt><b>Logentry Id</b></td>
	    <td class=alt><b>Incid. Emne</b><br />
	    Beskrivelse (fra logentry)
	    <hr style="color:#8caae6; border:0px #000000 solid; height:1px;" />Kommentar til. timereg.</td>
	    <td>
            &nbsp;</td>
	    
	</tr>
	
	
	<%
	
	ekspTxt = "Id;Status;Overført;Medarbejder;Init.;Starttid;Sluttid;Timer:Min;Kunde;job;Aktivitet;Incident;Log Entry;Emne;Beskrivelse;Kommentar;"
	ekspTxt = ekspTxt &"xx99123sy#z"
	
	if medid <> 0 then
	medKri = " s.medid = "& medid
	else
	medKri = " s.medid <> 0 "
	end if
	
	if jobid <> 0 then
	jobKri = " AND s.jobid = " & jobid
	else
	jobKri = ""
	end if
	
	totaltid = 0
	lastDato = "1/1/2002"
	
	stdatoKri = strAar &"/"& strMrd &"/"& strDag
	sldatoKri = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	if brugPer = "1" then
	perSQLKri = " AND s.sttid BETWEEN '"& stdatoKri &"' AND '"& sldatoKri &"'"
	else
	perSQLKri = ""
	end if 
	
	strSQL = "SELECT s.id AS sid, s.sttid, s.sltid, s.medid, s.jobid, s.aktid, s.incident, s.incidentlog, s.kommentar, "_
	&" j.jobnavn, jobnr, a.navn AS aktnavn, i.emne AS sdskemne, ilog.besk AS logbesk,"_
	&" kkundenavn, kkundenr, s.sid_godkendt, timereg_overfort, m.mnavn, m.mnr, m.init "_
	&" FROM stopur s "_
	&" LEFT JOIN job j ON (j.id = s.jobid) "_
	&" LEFT JOIN aktiviteter a ON (a.id = s.aktid) "_
	&" LEFT JOIN SDSK i ON (i.id = s.incident)"_
	&" LEFT JOIN sdsk_rel ilog ON (ilog.id = s.incidentlog)"_
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m ON (m.mid = s.medid)"_
	&" WHERE "& medKri &""& jobKri &""& perSQLKri &""& visSQL &""_
	&" ORDER BY "& orderBySQL &" DESC"
	
	'Response.Write strSQL
	'Response.flush
	
	t = 0
	g = 0
	f = 0
	oRec.open strSQL, oCOnn, 3
	While not oRec.EOF
	
	'if len(trim(oRec("sltid"))) <> 0 then
	'hg0_CHK = ""
	'hg1_CHK = "CHECKED"
	'else
	'hg1_CHK = ""
	'hg0_CHK = "CHECKED"
	'end if 
	
	
	'*** <br /><=incid % = <=oRec("incident")% **'
	
	'if (cint(incid) = oRec("incident") AND cint(incid)) <> 0 OR (cint(aktid) = oRec("aktid") AND cint(aktid) <> 0) then 
	
	'bghtis = "#FFFF99"
	'borderC = "orange"
	
	'else
	
	select case right(t,1)
	case 1,3,5,7,9
	bghtis = "#FFFFFF"
	case else
	bghtis = "#EFF3FF"
	end select
	
	borderC = "silver"
	
	'end if
  
	
	
	if len(trim(oRec("sttid"))) <> 0 then
	thisDato = datepart("ww", cdate(oRec("sttid")), 2,2)
	else
	thisDato = lastDato
	end if
 	
	if cdate(thisDato) <> cdate(lastDato) then 
	%>
	<tr><td colspan=12 height=30 bgcolor="#ffffe1" style="border-bottom:2px orange solid; padding-left:10px;">
	Uge: <b><%=thisDato %>, <%=year(oRec("sttid")) %></b>
	</td></tr>
	<%
	end if
	%>
	<form action="stopur_2008.asp?func=opd&logentry=<%=logentry%>&incid=<%=incid%>" method="POST" name="stopur" id="stopur">
	<tr bgcolor="<%=bghtis %>">
        <input id="antalsids_<%=oRec("sid") %>" name="antalsids" type="hidden" value="<%=oRec("sid")%>" />
        <td style="border-bottom:1px <%=borderC%> dashed; padding:0px 0px 0px 5px;"><%=oRec("sid")%>
        </td>
	    <td style="border-bottom:1px <%=borderC%> dashed; padding:5px 0px 0px 10px;"> 
	    <%
	    ekspTxt = ekspTxt & oRec("sid") &";" & oRec("sid_godkendt") &";" & oRec("timereg_overfort") &";"  
	    
	    if oRec("sid_godkendt") = 1 then
	    
	        'if oRec("timereg_overfort") <> 1 then 
	        'antalsids_gk = antalsids_gk & ", "& oRec("sid")
	        'img = "<IMG SRC='../ill/stopur_klar.png' alt='Med på liste der overføres til timeregistrering' />"
	        'g = g + 1
	        'else
	        f = f + 1
	        img = "<IMG SRC='../ill/stopur_eroverfort.png' alt='Er overført til timeregistrering' />"
	        'end if
	        
	    else
	    
	    
	        if oRec("jobid") <> 0 AND isDate(oRec("sttid")) AND isDate(oRec("sltid")) then%>
            <input id="sid_godkendt_<%=t%>" name="sid_godkendt_<%=oRec("sid")%>" value="1" CHECKED type="checkbox"/>
            <!--onclick="subm('stopur_2008.asp?func=opd&logentry=<%=logentry%>&incid=<%=incid%>&sid_godkendt_<%=oRec("sid")%>=1&FM_sttid_<%=oRec("sid") %>')"-->
            <%
            antalsids_gk = antalsids_gk & ", "& oRec("sid")
            g = g + 1
            
            img = "<input type=image src='../ill/stopur_st_stop.png'/>"
            else 
            
            'img = "<IMG SRC='../ill/start_stopur.png' alt='Igangværende stopur's entry' />"
            img = "<input type=image src='../ill/stopur_st_stop.png'/>"
            %>
            <input id="sid_godkendt_<%=t%>" name="sid_godkendt_<%=oRec("sid") %>" value="0" type="hidden" />
            <%end if %>
            
       <%end if%>
       
            <%=img %>
            &nbsp;
            </td>
	    <td style="border-bottom:1px <%=borderC%> dashed;" class=lille>
	    <b>
	    <%if len(oRec("mnavn")) > 15 then %>
	    <%=left(oRec("mnavn"), 15) %>..
	    <%else %>
	    <%=oRec("mnavn")%>
	    <%end if %> (<%=oRec("mnr") %>)</b><br />
	    
	    
	    <%if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
	    <input type="text" name="FM_sttid_<%=oRec("sid") %>" id="FM_sttid_<%=oRec("sid") %>" value="<%=oRec("sttid")%>" style="width:110px; font-size:9px;">
	    <%else %>
	    <input type="hidden" name="FM_sttid_<%=oRec("sid") %>" id="FM_sttid_<%=oRec("sid") %>" value="<%=oRec("sttid")%>">
	    <%=oRec("sttid")%>
	    <%end if %>
	    
	   </td>
	   
	    <%
	    ekspTxt = ekspTxt & oRec("mnavn") &" (" & oRec("mnr") &");"& oRec("init") &";" 
	    ekspTxt = ekspTxt & oRec("sttid") &";" & oRec("sltid") & ";"   
	    %>
	   
	    <td style="border-bottom:1px <%=borderC%> dashed; padding-top:15px;" class=lille>
	    <%if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
	    <input type="text" name="FM_sltid_<%=oRec("sid") %>" id="FM_sltid_<%=oRec("sid") %>" value="<%=oRec("sltid")%>" style="width:110px; font-size:9px;">
	    <%else %>
	    <input type="hidden" name="FM_sltid_<%=oRec("sid") %>" id="FM_sltid_<%=oRec("sid") %>" value="<%=oRec("sltid")%>">
	    <%=oRec("sltid")%>
	    <%end if %>
	    </td>
	    <td style="border-bottom:1px <%=borderC%> dashed;" align=center>
	    <%if isDate(oRec("sttid")) AND isDate(oRec("sltid")) then 
	    tidThis = datediff("n", oRec("sttid"),oRec("sltid"), 2,2)
	    else
	    tidThis = 0 
	    end if
	    
	    totaltid = totaltid + tidThis
	    %>
	    <%
	    call timerogminutberegning(tidThis)%>
	    <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	    </td>
	    
	    <%ekspTxt = ekspTxt & thoursTot &":"& left(tminTot, 2)  &";" %>
	    
	    <input name="hentgem_<%=oRec("sid") %>" id="Hidden2" value="0" type="hidden"  />
            
	    <!--
	    <td style="border-bottom:1px <%=borderC%> dashed;">
	        <%if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
            <input name="hentgem_<%=oRec("sid") %>" id="Radio1" value="0" type="hidden"  />
            <%else %>
               &nbsp;
               <%end if %></td>
               -->
               
         <!--      
	    <td style="border-bottom:1px <%=borderC%> dashed;">
	        <%if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
            <input name="hentgem_<%=oRec("sid") %>" id="Radio1" value="1" type="hidden" <%=hg1_CHK %> />
            <%else %>
               &nbsp;
               <%end if %></td>
               -->
          
	    <td style="border-bottom:1px <%=borderC%> dashed; padding:0px 5px 5px 10px;" class=lille>
	    <%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)<br />
	    <%if len(oRec("jobnavn")) <> 0 then  %>
	    <b><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)</b>
	    <%end if %>
	    &nbsp;
	    
	    <br />
	    <%
	    strSQLakt = "SELECT a.id, a.navn, a.fakturerbar FROM aktiviteter a WHERE a.job = " & oRec("jobid") &" AND "& right(strSQLkri3, len(strSQLkri3) - 17) &" ORDER BY a.fase, a.sortorder, a.navn" 
	    'Response.Write strSQLakt
	    'Response.flush
	    
	    if oRec("jobid") <> 0 then 
	    
	    if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then%>
	    <select name="aktid" id="aktid" style="width:200px; font-size:9px;"><%
	    end if
	    
	    
	    strAktnavn = ""
	    oRec2.open strSQLakt, oConn, 3
	    while not oRec2.EOF
	    
	    call akttyper(oRec2("fakturerbar"), 1) 
	    
	    
	    if oRec("aktid") = oRec2("id") then 
	    aSel = "SELECTED"
	    strAktnavn = oRec2("navn") &" ("& akttypenavn &")"
	    else
	    aSel = ""
	    end if 
	    
	    if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
	    <option value="<%=oRec2("id") %>" <%=aSel %>><%=oRec2("navn") %> (<%=akttypenavn %>)</option>
	    <%end if
	    
	    oRec2.movenext
	    wend 
	    oRec2.close
	    %>
	    <%if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
	    </select>
	    <%else %>
	    <%=strAktnavn%>
	    <%end if %>
	    
	    <%else %>
            <input id="aktid" name="aktid" value="0" type="hidden" />
	    <%end if %>
	    </td>
	    
	    <%ekspTxt = ekspTxt & oRec("kkundenavn") &" ("& oRec("kkundenr")&");" & oRec("jobnavn") &" ("& oRec("jobnr") &");"& strAktnavn &";" %>
	    
	    <td style="border-bottom:1px <%=borderC%> dashed;" align=center><%=oRec("incident") %></td>
	    <td style="border-bottom:1px <%=borderC%> dashed;" align=center><%=oRec("incidentlog") %></td>
	    
	    
	    <%
	    ekspTxt = ekspTxt & oRec("incident") &";"& oRec("incidentlog") &";"
	    %>
	    
	    
	    <%if len(oRec("sdskemne")) > 30 then
	    strEmne = left(oRec("sdskemne"),30) &".."
	    else
	    strEmne = oRec("sdskemne")
	    end if %>
	    
	   
	    
	    <%if len(oRec("logbesk")) > 50 then 
	    strBesk = left(oRec("logbesk"), 50) &".."
	    else
	    strBesk = oRec("logbesk") 
	    end if
	    
	    
	    %>
	    
	    <td style="border-bottom:1px <%=borderC%> dashed; padding-top:5px;;" class=lille>
	    <b><%=strEmne  %></b>
        <%=strBesk %> 
        <%if oRec("sid_godkendt") <> 1 AND print <> "j" AND (len(strEmne) <> 0 OR len(strBesk) <> 0) then %>
        <input type=button value=">>" style="font-size:9px; font-family:arial;" onclick="kopiertxt('<%=t %>');">
        <br />
        <%end if %>
	    
	    <% 
	    
	    if len(trim(oRec("logbesk"))) <> 0 then
	    strBeskhidfld = replace(oRec("logbesk"), Chr(34), "''")
	    strBeskhidfld = replace(strBeskhidfld, Chr(39), "'")
	    end if
	    
	    if len(trim(oRec("sdskemne"))) <> 0 then
	    strEmnehidfld = replace(oRec("sdskemne"), Chr(34), "''")
	    strEmnehidfld = replace(strEmnehidfld, Chr(39), "'")
	    end if
	    
	    %>
	   
        <input id="FM_hidden_txt_<%=t%>" type="hidden" value="<%=strEmnehidfld%> <%=strBeskhidfld %>" />
        
          
	    <%if oRec("timereg_overfort") <> 1 AND oRec("sid_godkendt") <> 1 AND print <> "j" then %>
	    <textarea id="FM_besk_<%=t%>" name="FM_besk" cols="25" rows="3" style="font-size:9px;"><%=oRec("kommentar") %></textarea>
        <%else %>
            <%if len(oRec("kommentar")) <> 0 then %>
            <hr style="color:#8caae6; border:0px #000000 solid; height:1px;" /><%=oRec("kommentar") %>
            <%end if %>
        <%end if %>
         </td>
          
          <%
          
          sdskEmne = ""
          if len(oRec("sdskemne")) <> 0 then
          sdskEmne = replace(oRec("sdskemne"), Chr(34), "''")
          else
          sdskEmne = ""
          end if
         
          logbesk = ""
          if len(oRec("logbesk")) <> 0 then
          logbesk = replace(oRec("logbesk"), Chr(34), "''")
          else
          logbesk = ""
          end if
          
          strKomm = ""
          if len(oRec("kommentar")) <> 0 then
          strKomm = replace(oRec("kommentar"), Chr(34), "''")
          else
          strKomm = ""
          end if
                    
          ekspTxt = ekspTxt & sdskEmne &";"& logbesk &";"& strKomm &";"
          ekspTxt = ekspTxt &"xx99123sy#z"
          %>
            
	    <td style="border-bottom:1px <%=borderC%> dashed;">
	    <%if print <> "j" AND oRec("timereg_overfort") <> 1 then %>
	    
	        <%if oRec("sid_godkendt") <> 1 then %>
	        <a href="stopur_2008.asp?func=slet&sid=<%=oRec("sid") %>">
            <img src="../ill/slet_16.gif" alt="Slet" border="0"/></a>
            <%else %>
             <a href="stopur_2008.asp?func=fortryd&sid=<%=oRec("sid") %>&aktid=<%=aktid %>&incid=<%=incid %>">
            <img src="../ill/stopur_fortryd.png" alt="Fortryd godkendt" border="0"/></a>
            <%end if %>
            
         <%else %>
         &nbsp;
         <%end if %></td>
	</tr>
	</form>
	<%
	lastDato = thisDato
	t = t + 1
	antalsids = antalsids & ", "& oRec("sid")
	
	oRec.movenext
	wend
	oRec.close %>
	<tr>
	    <td colspan=4>Antal: <b><%=t%></b>, På liste til timereg.: <b><%=g %></b>, Overført: <b><%=f %></b>
            &nbsp;</td>
            <td align=center><% 
            call timerogminutberegning(totaltid)
            %>
            <b><%=thoursTot &":"& left(tminTot, 2) %></b>
            </td>
            <td colspan=7>
                &nbsp;</td>
	</tr>
	</table>
    </div>
    
	
	
	<%if print <> "j" then%>
	
	
	<table cellpadding=2 cellspacing=1 border=0 width=100%>
	<%if g <> 0 then%>
	<form action="timereg_akt_2006.asp?func=db&stopur=1" method="POST">
	<input id="antalsids_t" name="antalsids_t" type="hidden" value="<%=t%>" />
	<input id="sids" name="sids" type="hidden" value="<%=antalsids_gk%>" />
	<input id="medid" name="medid" type="hidden" value="<%=medid%>" />
	<tr>
	    <td align=right>
	    Overfør <IMG SRC='../ill/stopur_klar.png' alt='' /> markerede Stopur's Entries til timregistrering<br />
        <input id="Submit1" type="submit" value="Overfør >> " />
        </td></tr></form>
    <%else %>
    <tr><td align=right>Der er ingen <IMG SRC='../ill/stopur_klar.png' alt='' /> markerede Stopur's Entries klar til timeregistrering.</td></tr>
	<%end if %>
	</table>
	</div>
	
	
<br><br>
    <div id="sideinfo" style="position:relative; left:0px; top:20px; width:400px; visibility:visible; padding:10px; border:1px red dashed; background-color:#ffffe1;">
			<table>
			<tr><td>
			<img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
			
	* Kun Incindents der er tilknyttet et job kan overføres til timeregistrering.<br />
	* Tilknyt Incidents til et job, ved at klikke på et Incident på Incident listen.<br />
	* Kun godkendte (<i>V</i>) stopurs entrys kan overføres til timeregistrering.<br />
	* Indlæste timer kan altid ændres fra timereg. siden, så længe jobbet er åben for timeregistrering.
    <br /><br />
    <b>Signatur forklaring:</b><br />
    <table cellspacing=2 cellpadding=2 border=0 width=100%>
    <tr>
    <td width=50 style="padding-left:6px;"><IMG SRC="../ill/stopur_st_stop.png" alt="Igangværende stopur's entry" /></td><td>Igangværende stopur's entry - Start / Stop
    </tr><tr>
    <td><input id="Checkbox1" type="checkbox" /></td><td> Entry er klar til at blive tilføjet til den liste der skal overføres til timeregistrering
    </tr>
        <tr><td style="padding-left:6px;"><IMG SRC="../ill/stopur_klar.png" alt="Med på liste der overføres til timeregistrering" /></td><td> Med på liste over entries der overføres til timeregistrering 
    </tr>
    <tr><td style="padding-left:6px;"><IMG SRC="../ill/stopur_eroverfort.png" alt='Er overført til timeregistrering' /></td><td> Er overført til timeregistrering
    </td></tr></table>
    
    </td></tr></table>
			</div>
	<br><br>&nbsp;
	<%else %>
	<br /><br />&nbsp;
<%end if 
	
	
if print <> "j" then

ptop = 70
pleft = 700
pwdt = 120

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="stopur_2008_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=strDag&"/"&strMrd&"/"&strAar & " - " & strDag_slut&"/"&strMrd_slut&"/"&strAar_slut%>">
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=ekspTxt%>">
            <input type="hidden" name="txt20" id="txt20" value="">

<tr>
    <td align=center><input type="image" src="../ill/export1.png">
    </td>
    </td><td>.csv fil eksport</td>
    </tr>
    <tr>
    <td align=center><a href="stopur_2008.asp?print=j" target="_blank"  class='vmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
</td><td>Print version</td>
   </tr>
   
	</form>
   </table>
</div>
<%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if%>
    


</div>

<%
end select

end if 'user session
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





