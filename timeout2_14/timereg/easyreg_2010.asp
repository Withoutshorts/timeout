<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%

        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        'case "FM_ajaxstatus"
        
        case "FN_getCustDesc"
        
        if Request.Form("cust") <> 0 then
        sqlWh = " job = " & Request.Form("cust")
        else
        sqlWh = " job = 0"
        end if
        
        
        strSQLakrt = "SELECT a.navn AS aktnavn, id FROM "_
        &" aktiviteter AS a WHERE "& sqlWh &" ORDER BY navn"
        
        'Response.Write strSQLakrt
        'Response.end
        
        a = 0
        oRec.open strSQLakrt, oConn, 3
        while not oRec.EOF
        
         '*** ÆØÅ **'
         jq_format(oRec("aktnavn"))
         aktnavn = jq_formatTxt
        
        response.write "<option value='"& oRec("id")&", #' >"& aktnavn &"</option>"
        
        a = a + 1
        oRec.movenext
        wend
        oRec.close
        
        if a = 0 then
         Response.Write "<option value='0, #'>Der blev ikke fundet nogen aktiviteter.</option>" 
        end if
        
        end select
        Response.end
        end if  


%>

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato.asp"-->
<%

func = request("func")


if len(trim(request("FM_jobid"))) <> 0 then
jobid = request("FM_jobid")
else
jobid = 0
end if 

if func = "opd_easyreg" then
		
		useEasy = request("FM_use_easy")
		
		'Response.Write useEasy & "<br>"
		
		'if len(trim(useEasy)) <> 0 then
		'len_useEasy = len(useEasy)
		'left_useEasy = left(useEasy, (len_useEasy-3))
		'useEasy = left_useEasy
		'end if
		
		aktiverPUSH = 0
		intuseEasy = Split(useEasy, "#, ")
	    'intuseEasy = replace(intuseEasy, ",", "")
		
		'*** Nulstiller alle ***'
		strSQLeaN = "UPDATE aktiviteter SET easyreg = 0 WHERE id <> 0"
	   	oConn.execute(strSQLeaN)
		
		for j = 0 to UBOUND(intuseEasy)
		
		    if len(trim(intUseEasy(j))) > 1 then
	   	    intUseEasy(j) = trim(left(intUseEasy(j), len(intUseEasy(j)) - 2))
	   	    else
	   	    intUseEasy(j) = 0
	   	    end if
	   	    
	   	    intUseEasy(j) = replace(intUseEasy(j), ",", "")
	   	    intUseEasy(j) = replace(intUseEasy(j), "#", "")
	   	    intUseEasy(j) = trim(intUseEasy(j))
	   	    
	   	    strSQLea = "UPDATE aktiviteter SET easyreg = 1 WHERE id = "& intUseEasy(j)
	   	    'Response.Write strSQLea & "<br>"
	   	    'Response.flush
	   	    oConn.execute(strSQLea)
	   	    
	   	    '*** efter NUL tilføjes PUSH dvs det er en ny aktivitet 
	   	    '*** der tilføejs easyreg listen og der skal den stå aktiv 
	   	    '*** for alle medarbejdere
	   	   
	   	    
	   	    if cint(aktiverPUSH) = 1 then
	   	    
	   	    strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & intUseEasy(j) & " WHERE jobid = " & jobid
	   	    'Response.Write strSQLtreguse & "<br>"
	   	    'Response.flush
	   	    oConn.execute(strSQLtreguse)
	   	    
	   	    end if
	   	    
	   	    if intUseEasy(j) = 0 then
	   	    aktiverPUSH = 1
	   	    end if
	   	
	   	next
		
		'Response.Write useEasy
		'Response.end
		
		
		
		Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.location.href('easyreg_2010.asp');</script>")
		
		'Response.Redirect "easyreg_2010.asp?FM_jobid="&jobid
		

else

%>
<script src="inc/easyreg_jav.js">
    // JScript File

    

</script>



<div id="sindhold" style="position:absolute; left:20; top:20; width:60%; height:600; visibility:visible;">
<%

'Response.write usemrn
'Response.flush

	
	'************************************************************************************************
	'**** Guiden... Vis Kunder ********
	'************************************************************************************************
	%>
	
	
	
	<h4>Den centrale Easyreg. adminstration</h4>
	
	Her er en oversigt over alle de aktiviteter der er med på den aktuelle Easyreg. liste.
	Hver medarbejder har kun adgang til de aktiviteter, som medarbjederen er tilknyttet via sine projektgrupper, og som er aktive på deres personlige <b>aktivliste</b>.
	<br>
	<br>
    &nbsp;</td>
	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
		
	<table cellspacing="0" cellpadding="0" border="0" width=700>
	<form action="easyreg_2010.asp?func=opd_easyreg" method="post" name="usejobguide">
 	<tr bgcolor="#5582d2">
		<td class=alt style="padding:5px; width:300px;"><b>Job</b></td>
		<td style="height:25px; padding:5px;" class=alt><b>Aktivitet</b></td>
        <td style="height:25px; padding:5px;" class=alt>Fase</td>
		<td class=alt style="padding:5px;">&nbsp;</td>
		
	</tr>
	<%
	
	'call hentbgrppamedarb(usemrn)
	
	lastKid = 0
	strSQL = "SELECT j.jobstatus, j.jobnavn, j.jobnr, j.jobknr, "_
	&" Kid, Kkundenavn, Kkundenr, a.id, a.navn AS aktnavn, a.fakturerbar, a.easyreg, a.fase "_
	&" FROM aktiviteter AS a "_
	&" LEFT JOIN job AS j ON (j.id = a.job) "_
	&" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
	&" WHERE j.jobstatus = 1 AND a.aktstatus = 1 AND a.easyreg = 1 "_
	&" ORDER BY Kkundenavn, jobnavn, a.fase, a.navn"
	
	'Response.write (strSQL)	
	'Response.flush
	x = 0
    antalSelEasy = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
			
		if lastKid <> oRec("Kid") then
		%>
		<tr>
			<td colspan="5" bgcolor="#8caae6" height=30 class=alt style="border-bottom:1px #999999 solid;">&nbsp;&nbsp;<b><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr")%>)</b></td>
		</tr>
		<%else%>
		<tr>
			<td bgcolor="#999999" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<%end if
		
		'**** Er jobbet allerede valgt?
		'strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &" AND jobid = " & oRec("id")
		'oRec3.open strSQL3, oConn, 3
		
		
		
		if oRec("easyreg") <> 0 then
		selEasy = "CHECKED"
        antalSelEasy = antalSelEasy + 1 
		else
		selEasy = ""
		end if
		
		select case right(x, 1)
		case 0,2,4,6,8
		trbg = "#FFFFFF"
		case else
		trbg = "#D6Dff5"
		end select
		
		%>
		<tr bgcolor="<%=trbg%>">
			<td height=20 valign=top style="padding:4px 10px 2px 2px;"><b><%=oRec("jobnavn")%></b> (<%=oRec("jobnr")%>)</td>
			<td valign=top style="padding:4px 10px 2px 2px;"><%=oRec("aktnavn") %></td>
            <td valign=top style="padding:4px 10px 2px 2px;"><span style="color:#999999;"><%=oRec("fase") %></span></td>
			<td valign=top style="padding:4px 10px 2px 2px;"><input type="checkbox" class="alljobCHK" name="FM_use_easy" value="<%=oRec("id")%>" <%=selEasy%>></td>
			<td>
			
			</td>
            
            <input id="Hidden2" name="FM_use_easy" value="#" type="hidden" />
			</tr>
		<%
		
		'Response.flush
		lastKid = oRec("Kid")
		x = x + 1
		oRec.movenext
		wend
		oRec.close
		%>
		
	
			<input id="Hidden5" name="FM_use_easy" value="0, #" type="hidden" />
	
		<!--
		<tr bgcolor="#D6DFF5">
			<td colspan="4" align="right"><br>Vis denne guide næste gang du logger på?&nbsp;Ja:<input type="checkbox" name="FM_visguide" value="j"></td>
		</tr>-->
		<tr>
			<td style="border-top:1px #003399 solid;" colspan="6" align="right"><br>
			
                <input id="Submit1" type="submit" value="<%=tsa_txt_144 %> >>" /></td>
		</tr>
		</table>

        Der er <b><%=antalSelEasy %> Easyreg. aktiviteter.</b>
		
		</div>
		
		<h3>Find aktiviteter til Easyreg. listen</h3>
		Ny tilføjede aktiviteter bliver automatisk gjort aktive på medarbejdernes Easyreg. liste. <b>(gælder alle Easyreg. aktiviteter på jobbet)</b>
		Eksisterende Easyreg. aktivitetrer beholder deres indstillinger på den enkelte medarbejers <b>aktivliste.</b>
		<br /><br />
		
		
		<%if antalSelEasy < 150 then %>
		
		<table cellspacing=0 cellpadding=0 border=0 width=600>
	
		<tr><td valign=top><b>Job:</b></td><td>
		   
		    <%strSQLjob = "SELECT j.id AS jid, jobnavn, jobnr, jobknr, kkundenavn, kkundenr, kid, COUNT(a.id) AS antal_akt"_
		    &" FROM job AS j "_
		    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
		    &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND aktstatus = 1) "_
		    &" WHERE j.jobstatus = 1 AND kkundenavn <> '' GROUP BY job ORDER BY kkundenavn, jobnavn"
		    
		    'Response.Write strSQLjob
		    'Response.flush
		    
		    
		   
		    
		    %> 
		    <select id="FM_jobid" name="FM_jobid" style="width:400px; font-size:10px;" size=10><%
		    x = 0
		    oRec.open strSQLjob, oConn, 3
		    while not oRec.EOF
		    
		    if lastKid <> oRec("kid") OR x = 0 then
		        if x <> 0 then%>
		        <option disabled></option>
		        <%end if %>
		    <option disabled></option>
            <option disabled><%=oRec("kkundenavn") & " ("& oRec("kkundenr") &")" %> </option>
		    <%end if
		    
		    
		    if jobid <> 0 then
		        if cint(jobid) = cint(oRec("jid")) then
		        jSEl = "SELECTED"
		        else
		        jSEl = ""
		        end if
		    else
		        
		        if x = 0 then
		        jSEl = "SELECTED"
		        else
		        jSEl = ""
		        end if
		    end if
		    
		    %>
		    <option value="<%=oRec("jid")%>" <%=jSEl %>><%=oRec("jobnavn") & " ("& oRec("jobnr") &") - "& oRec("antal_akt") & " akt." %></option>
		    <%
		    lastKid = oRec("kid")
		    x = x + 1
		    oRec.movenext
		    wend 
		    oRec.close
		    
	        %>
		    </select>
		    
            </td></tr>
            
        
		     
            
            
      
      
		<tr><td valign=top><br /><b>Aktiviterer:</b><br />
		Vælg gerne flere</td>
		<td>
		<br>
            <!--<textarea id="fajl" cols="60" rows="10"></textarea>-->
		
		<select MULTIPLE style="width:400px; font-size:10px;" size=5 name="FM_use_easy" id="jq_aktlist">
	
		<!-- jquery Akt val -->
		</select></td></tr>
		<tr><td colspan=2 align=right>
            <input id="Submit2" type="submit" value=" Tilføj >> " /></td></tr>
        
		</form>
		</table>

        <%else

        itop = 180
ileft = 185
iwdt = 300
idsp = ""
ivzb = "visible"
iId = "tregaktmsg1"
call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
%>
			
            <b>Der er for mange Easyreg. aktiviteter</b><br><br />
	        Der er mere end 150 Easyreg. aktiviteter og der kan derfor ikke tilføjes flere.<br><br>
	        
	        
	</td></tr></table>
	</div>
	

<%
end if
		
      %>
	<br><br>&nbsp;
	
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->

<%end if%>
