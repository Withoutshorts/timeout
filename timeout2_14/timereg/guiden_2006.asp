<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato.asp"-->

<%
usemrn = request("mid")
func = request("func")
lc = request("lc")

if func = "updguide" then
		
		useJob = request("FM_use_job")
		useAkt = request("FM_use_akt")
        useEasy = request("FM_use_easy")
		
        
		
		'delicounter = 1
		'Response.Write useJob & "<hr>"
		'Response.Write useAkt & "<br>"
		'Response.end
		'forvalgt = 0 'off / 1 = on
        forvalgt = request("FM_useaktive_job")

		call setGuidenUsejob(usemrn, useJob, useAkt, 0, useEasy, forvalgt, 0)
		
		
		
		'Response.end
		
		'**** Bruges ikke mere timreg_2006 ***
		'if request("FM_visguide") = "j" then
			oConn.execute("UPDATE medarbejdere SET visguide = 1 WHERE Mid = "& usemrn &"")
			'else
			'oConn.execute("UPDATE medarbejdere SET visguide = 0 WHERE Mid = "& usemrn &"")
		'end if
		
		'Response.end
		
		if lc = "res" then
		'Response.Write("<script language=""JavaScript"">window.opener.location.href('timereg_akt_2006.asp?hideallbut_first=1');</script>")
        'Response.Write("<script language=""JavaScript"">window.opener.location.href('ressource_belaeg_jbpla.asp?menu=webblik');</script>")
        Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")
		else
        Response.Write("<script language=""JavaScript"">window.opener.location.href('timereg_akt_2006.asp?menu=treg');</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")
		
		'Response.Write("<script language=""JavaScript"">window.opener.top.frames['t'].location.reload();</script>")
		'Response.Write("<script language=""JavaScript"">window.opener.top.frames['a'].location.reload();</script>")
		'Response.Write("<script language=""JavaScript"">window.close();</script>")
		end if 

else

%>
<script>
    // JScript File

    $(document).ready(function () {

        $("#allejob").click(function () {

            if ($("#allejob").is(':checked') == true) {
                $(".alljobCHK").attr('checked', true);
            } else {
                $(".alljobCHK").attr('checked', false);

            }

        });

        $("#alleeasy").click(function () {

            if ($("#alleeasy").is(':checked') == true) {
                $(".alleasyCHK").attr('checked', true);
            } else {
                $(".alleasyCHK").attr('checked', false);

            }

        });

        $("#alleaktivejob").click(function () {

            //alert("her")

            if ($("#alleaktivejob").is(':checked') == true) {
                $(".alleaktivejobCHK").attr('checked', true);
            } else {
                $(".alleaktivejobCHK").attr('checked', false);

            }

        });


        $(".all_akt_job").click(function () {

            var thisid = this.id

            var idlngt = thisid.length
            var idtrim = thisid.slice(12, idlngt)


            if ($("#all_akt_job_" + idtrim).is(':checked') == true) {
                $(".all_akt_job_" + idtrim).attr('checked', true);
            } else {
                $(".all_akt_job_" + idtrim).attr('checked', false);

            }


        });

        $(".alleaktivejobCHK").click(function () {

            //alert("344")
            var thisid = this.id

            var idlngt = thisid.length
            var idtrim = thisid.slice(12, idlngt)


            // Skal kun slå fra
            if ($("#alx_akt_job_" + idtrim).is(':checked') == true) {
                //$(".all_akt_job_" + idtrim).attr('checked', true);
                $(".all_akt_job").attr('checked', false);
                $("#bank_akt_job_" + idtrim).attr('checked', true);
            } else {
                $(".all_akt_job_" + idtrim).attr('checked', false);
                $(".all_akt_job").attr('checked', false);

            }


        });




        //$(".sp_showakt").mouseover(function () {
        //$(this).css('cursor', 'pointer');
        //});

        //$(".tr_a_0").css('visibility', 'hidden')
        //$(".tr_a_0").css('display', 'none')

        

    });


    function tjekstatus(aid, jid) {
        
        thisAid = document.getElementById(aid)
        if (thisAid.checked == 1) {
            //alert(aid)
            document.getElementById("alx_akt_job_" + jid + "").checked = 1;
            document.getElementById("bank_akt_job_" + jid + "").checked = 1;
        }
       
    }

</script>

<!--<SCRIPT language=javascript src="inc/timereg_2006_func.js"></script>-->

<div id="sindhold" style="position:absolute; left:20; top:20; width:60%; height:600; visibility:visible;">
<%

call positiv_aktivering_akt_fn()
call showEasyreg_fn()
call meStamdata(usemrn)
            
'Response.write usemrn
'Response.flush

if cint(positiv_aktivering_akt_val) = 1 then
jobWdt = "250"
else
jobWdt = "400"
end if

	'************************************************************************************************
	'**** Guiden... Vis Kunder ********
	'************************************************************************************************
	%>
	
	
	
	<h4><%=tsa_txt_142 %></h4>
	<%=tsa_txt_143 %>
	<br><br />
	Vælg indstillinger for: <b><%=meTxt %></b><br /><br />
    
	
	Søg [CTRL] + [F]	
	<table cellspacing="0" cellpadding="2" border="0" width="800">
	<form action="guiden_2006.asp?func=updguide&mid=<%=usemrn%>&lc=<%=lc%>" method="post" name="usejobguide">
 	<tr bgcolor="#5582d2">
		<td class=alt style="padding:5px;" valign=top><b>Job</b><br />
		</td>
		<td style="height:25px; padding:5px;" valign=top class=alt>
            <input type="checkbox" id="allejob" value="1"/><b>Aktiver job i søgefeltet (job-banken)</b><br /> 
            <span style="font-size:9px; color:#cccccc;">Job du er tilknyttet via dine projektgrupper.</span>
           
		</td>
       <td class=alt style="padding:5px;" valign=top> <input type="checkbox" id="alleaktivejob" value="1"/> <b>Personlig aktiv jobliste.</b><br />
        <span style="font-size:9px; color:#cccccc;">Job som du arbejder på netop nu, og som er med på din aktive jobliste.</span></td>

        <%if cint(showEasyreg_val) = 1 then  %>
        	<td class=alt style="padding:5px;" valign=top> <input type="checkbox" id="alleeasy" value="1"/><b>Easyreg.</b> (antal) aktiviteter<br />
		 <span style="font-size:9px; color:#cccccc;">Vises på Easyreg. listen (hvis job er slået til og aktivitet er valgt som Easyreg. aktivitet)</span></td>
        <%else %>
        <td>&nbsp;</td>
        <%end if %>

        <%if cint(positiv_aktivering_akt_val) = 1 then %>
		<td class=alt style="padding:5px;" valign=top><b>Viser kun de valgte aktiviteter</b><br />(Positiv tildeling slået til) <!--<input type="checkbox" id="allall_job_akt" value="1"/><b>Aktiviteter</b><br />
        <span style="font-size:9px; color:#cccccc;">Aktiviteter du er tilknyttet via dine projektgrupper.</span>
        -->&nbsp;</td>
        <%else %>
		<td>&nbsp;</td>
        <%end if %>
	</tr>
	<%
	
	call hentbgrppamedarb(usemrn)
	
    '*** Søger efter job først, da timereg_usejob kun indeholder de job der er aktive lige nu. 
    '*** Du kan godt være tilmeldt et job via dine projektgrupper uden det er med i timereg_usejob

	lastKid = 0
	strSQL = "SELECT j.id, j.jobstatus, j.jobnavn, j.jobnr, j.jobknr, Kid, Kkundenavn, Kkundenr, "_
	&" useasfak, g.jobid AS gjid, forvalgt, g.easyreg, COUNT(a.id) AS antal_aeasy, risiko"_
    &" FROM job j "_
    &" LEFT JOIN kunder ON (kunder.Kid = j.jobknr)"_
	&" LEFT JOIN timereg_usejob AS g ON (g.medarb = "& usemrn &" AND g.jobid = j.id) "_
    &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1 AND a.aktstatus = 1)"_
	&" WHERE (jobstatus = 1 OR jobstatus = 3) AND fakturerbart = 1 AND kunder.Kid = j.jobknr "& strPgrpSQLkri &""_
	&" GROUP BY j.id ORDER BY Kkundenavn, jobnavn"
	
    'Response.write (strSQL)	
	'Response.flush
    
    'antal_aeasyIalt = 0
    antalJobIalt = 0
	x = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
			
		
		
		'**** Er jobbet allerede valgt?
		'strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &" AND jobid = " & oRec("id")
		'oRec3.open strSQL3, oConn, 3
		
		if len(trim(oRec("gjid"))) <> 0 then
		selJob = "CHECKED"
        antalJobIalt = antalJobIalt + 1
		else
		selJob = ""
		end if

        if oRec("forvalgt") <> 0 then
		selAktiveJob = "CHECKED"
        else
		selAktiveJob = ""
		end if
		
        
		if oRec("easyreg") <> -1 then
		selEasy = "CHECKED"
		else
		selEasy = ""
		end if

		select case right(x, 1)
		case 0,2,4,6,8
		trbg = "#FFFFFF"
		case else
		trbg = "#F7F7F7"
		end select
		
		%>
        
		<tr bgcolor="<%=trbg%>">
			<td height=30 style="padding:2px 5px 2px 5px; white-space:nowrap; width:<%=jobWdt%>px; border-top:1px #CCCCCC solid;" valign="top"><span style="font-size:9px;"><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr")%>)</span><br />
            
            
            <b><%=oRec("jobnavn")%></b> (<%=oRec("jobnr")%>)</td>
			<td valign="top" style="border-top:1px #CCCCCC solid; width:50px;"><input type="checkbox" class="alljobCHK" name="FM_use_job" id="bank_akt_job_<%=oRec("id")%>" value="<%=oRec("id")%>" <%=selJob%>></td>
            <td valign="top" style="border-top:1px #CCCCCC solid; width:50px;"><input type="checkbox" class="alleaktivejobCHK" name="FM_useaktive_job" id="alx_akt_job_<%=oRec("id")%>" value="1" <%=selAktiveJob%>></td>

             <%if cint(showEasyreg_val) = 1 then  %>
            <td valign="top" style="border-top:1px #CCCCCC solid; width:50px;">
			<input type="checkbox" class="alleasyCHK" name="FM_use_easy" value="<%=oRec("id")%>" <%=selEasy%>>
			
			
			    <%if oRec("antal_aeasy") <> "0" AND oRec("easyreg") <> 0 AND isNull(oRec("antal_aeasy")) = False then %>
			    <%=oRec("antal_aeasy") %> stk. Easyreg. aktiviteter
            
            
			    <%
                if oRec("easyreg") <> 0 AND len(trim(oRec("antal_aeasy"))) > 0 then
                antalEA = cint(oRec("antal_aeasy"))/1
                antal_aeasyIalt = antal_aeasyIalt + antalEA
                end if
            
                else %>
			    <%end if %>
			</td>
            <%else %>
            <td style="border-top:1px #CCCCCC solid;">&nbsp;</td>
            <%end if %>

			<%if cint(positiv_aktivering_akt_val) = 1 AND oRec("risiko") >= 0 then  %>
            <td valign="top" style="border-top:1px #CCCCCC solid; width:200px;"><input type="checkbox" class="all_akt_job" id="all_akt_job_<%=oRec("id")%>" name="FM_alljob_akt" value="0" onclick="tjekstatus('all_akt_job_<%=oRec("id")%>', '<%=oRec("id")%>');"><span style="color:#999999;">alle aktiviteter på job</span></td>
            <%else %>
            <td style="border-top:1px #CCCCCC solid;">&nbsp;</td>
            <%end if %>
            <input id="Hidden2" name="FM_use_job" value="#" type="hidden" />
            <input id="Hidden3" name="FM_useaktive_job" value="#" type="hidden" />
            <input id="Hidden11" name="FM_use_easy" value="#" type="hidden" />
            <input id="Text1" name="FM_use_akt" value="0" type="hidden" />
            <input id="Text2" name="FM_use_akt" value="#" type="hidden" />
			</tr>

            <%if cint(positiv_aktivering_akt_val) = 1 AND oRec("risiko") >= 0 then  %>
            <tr class="tr_a_0" bgcolor="<%=trbg%>" style="visibility:visible; display:;">
            <td colspan=4>&nbsp;</td>
            <td style="white-space:nowrap;">
		
			
            <%strSQLa = "SELECT a.navn AS aktnavn, a.fase, a.id, g.aktid AS gaktid FROM aktiviteter AS a "_
            &" LEFT JOIN timereg_usejob AS g ON (g.medarb = "& usemrn &" AND g.jobid = a.job AND g.aktid = a.id) "_ 
            &" WHERE a.job = "& oRec("id") & " AND a.aktstatus = 1 "& replace(strPgrpSQLkri, "j.", "a.") &" GROUP BY a.id ORDER BY a.fase, a.navn"
            'Response.write strSQLa
            'Response.flush
            oRec3.open strSQLa, oConn, 3
            While not oRec3.EOF 

            if len(trim(oRec3("gaktid"))) <> 0 then
		    selAkt = "CHECKED"
            else
		    selAkt = ""
		    end if
            
            %>
            	<input type="checkbox" class="all_akt_job_<%=oRec("id")%>" id="akt_<%=oRec3("id")%>" name="FM_use_akt" value="<%=oRec3("id")%>" <%=selAkt%> onclick="tjekstatus('akt_<%=oRec3("id")%>', '<%=oRec("id")%>');"> <%=oRec3("aktnavn") %> 
                <%if isnull(oRec3("fase")) <> true AND len(trim(oRec3("fase"))) <> 0 then %>
                <span style="font-size:9px; color:#999999;">fase: <%=oRec3("fase") %></span>
                <%end if %>
                
                <br />
            <input id="Hidden10" name="FM_use_job" value="<%=oRec("id")%>" type="hidden" />
            <input id="Hidden1" name="FM_use_akt" value="#" type="hidden" />
            <input id="Hidden7" name="FM_use_job" value="#" type="hidden" />
            <input id="Hidden13" name="FM_use_easy" value="0" type="hidden" /><!-- styres på job og fra Easyreglisten, bare så array passer-->
            <input id="Hidden12" name="FM_use_easy" value="#" type="hidden" />
            <input id="Hidden8" name="FM_useaktive_job" value="1" type="hidden" />
            <input id="Hidden9" name="FM_useaktive_job" value="#" type="hidden" />
            <%
            oRec3.movenext
            wend
            oRec3.close %>
			
			
			</td>
        </tr>
        <%end if %>
		<%
		
		'Response.flush
		lastKid = oRec("Kid")
		x = x + 1
		oRec.movenext
		wend
		oRec.close
		%>
		
		<input id="Hidden4" name="FM_use_job" value="" type="hidden" />
	    <input id="Hidden6" name="FM_useaktive_job" value="" type="hidden" />
	    <input id="Hidden5" name="FM_use_akt" value="" type="hidden" />
        <input id="Hidden14" name="FM_use_easy" value="" type="hidden" />
		<!--
		<tr bgcolor="#D6DFF5">
			<td colspan="4" align="right"><br>Vis denne guide næste gang du logger på?&nbsp;Ja:<input type="checkbox" name="FM_visguide" value="j"></td>
		</tr>-->
		<tr>
			<td style="border-top:1px #CCCCCC solid;" colspan="4" align="right"><br>
			
                <input id="Submit1" type="submit" value="<%=tsa_txt_144 %>" /></td>
		</tr>
		</form>
		</table>
        Antal job slået til: <%=antalJobIalt %>
        <!--<br />
        Antal Easyreg. aktiviteter slået til: <%=antal_aeasyIalt %>
        -->
	<br><br>&nbsp;
	
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->

<%end if%>
