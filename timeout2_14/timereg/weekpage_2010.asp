<%response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/smiley_inc.asp"-->



<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
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
	thisfile = "weekpage_2010.asp"

    rdir = request("rdir")
	
	call erStempelurOn()
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	'sendemail = request("sendemail")
	
	
	medarbid = request("medarbid")
	stDatoSQL = year(request("st_dato")) & "-" & month(request("st_dato")) & "-" & day(request("st_dato"))
	
	slDatoberegn = dateadd("d", 6, request("st_dato"))
	slDatoSQL = year(slDatoberegn) & "-" & month(slDatoberegn) & "-" & day(slDatoberegn)
	visning = 0
	
	select case func 
	case "-"
	
    case "opdaterstatus"

          ujid = split(request("ids"), ",")
	
    for u = 0 to UBOUND(ujid)
	
	            editor = session("user")


                if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                else
                uGodkendt = 0
                end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u)
				
				'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

    Response.Redirect "weekpage_2010.asp?medarbid="& medarbid &"&st_dato="& request("st_dato") &"&func=us"

	case "godkendugeseddel"
	
	 
	    strSQLup = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tmnr = "& medarbid & " AND tdato BETWEEN '"& stDatoSQL &"' AND '" & slDatoSQL & "'" 
	    oConn.execute(strSQLup)


        '*** Godkend uge status ****'
        call godekendugeseddel(thisfile, session("mid"), medarbid, stDatoSQL)
       

	    
	Response.Redirect "weekpage_2010.asp?medarbid="& medarbid &"&st_dato="& request("st_dato") &"&func=us"
	
	
    case "afvisugeseddel"
	
	'*** Afviser ugeseddel ***'
	if len(trim(request("FM_afvis_grund"))) <> 0 then
    txt = replace(request("FM_afvis_grund"), "'", "")
    else
    txt = ""
    end if

    
    call afviseugeseddel(thisfile, session("mid"), medarbid, stDatoSQL, slDatoSQL, txt)
	    
	Response.Redirect "weekpage_2010.asp?medarbid="& medarbid &"&st_dato="& request("st_dato") &"&func=us"

    case "adviserugeafslutning"

    call afslutugereminder(thisfile, session("mid"), medarbid, stDatoSQL, slDatoSQL, txt)
	    
	Response.Redirect "weekpage_2010.asp?medarbid="& medarbid &"&st_dato="& request("st_dato") &"&func=us&showadviseringmsg=1"
	
	case else
	
	
	
    
	
	
	
	call akttyper2009(2)
	
	thisweek = day(stdatoSQL) & "-" & month(stdatoSQL) &"-"& year(stdatoSQL)

    'Response.Write thisweek

    nextWeek = dateadd("d", 7, thisweek)
    prevWeek = dateadd("d", -7, thisweek)

    level = session("rettigheder")

    call ersmileyaktiv()
	
	leftPos = 20
	topPos = 20
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<script>
	    $(document).ready(function() {
	    $("#udspec").click(function() {

	        if ($("#udspecdiv").css('display') == "none") {

	            $("#udspecdiv").css("display", "");
	            $("#udspecdiv").css("visibility", "visible");
	            $("#udspecdiv").show(4000);

	        } else {
	            $("#udspecdiv").hide(1000);
	        }
	    });
	});
	</script>
	
	
	
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	
	<!-- ugeseddel -->
   <div id="afstem" style="display:; visibility:visible; width:900px; height:2020px; overflow:auto; background-color:#FFFFFF; padding:2px; border:1px #8caae6 solid; z-index:2000;">
    
    <%if func = "us" then 
        
    showtotal = 1
    showheader = 1%>
    <%call ugeseddel(medarbid,stdatoSQL,sldatoSQL,visning,showtotal,showheader)  %>
    
    <%else %>
    
    
    <!-- logind historik -->
    <table cellpadding=0 cellspacing=0 border=0 width=100%><tr>
     <td valign=top width=70% style="padding:10px;">
     
	<%call fLonTimerPer(stdatoSQL, 7, 0, medarbid) %>
	
	
	
	</td>
	<td align=right valign=top style="padding:10px 10px 0px 0px;">
	<table cellpadding=0 cellspacing=0 border=0 width=80>
	<tr>
	<td valign=top align=right><a href="weekpage_2010.asp?medarbid=<%=medarbid%>&st_Dato=<%=prevWeek%>&func=lo"><img src="../ill/nav_left_blue.png" border="0" /></a></td>
   <td style="pading-left:20px;" valign=top align=right><a href="weekpage_2010.asp?medarbid=<%=medarbid%>&st_Dato=<%=nextWeek%>&func=lo"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	</tr>
	</table>
	</td></tr>
	<tr>
	<td valign=top colspan=2 style="padding:10px;">
    
    
	<%
    d_end = 6
    call stempelurlist(medarbid, 0, 1, stdatoSQL, sldatoSQL,0, d_end, lnk) %>
	
	</td></tr>
    </table>
    
    
    </td></tr>
    </table>
    
    
    <%end if %>
   
   </div>
	
	    
	
	<br /><br /><br /><br /><br /><br /><br /><br />
        &nbsp;
			
			</div>
	
	<% 
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
