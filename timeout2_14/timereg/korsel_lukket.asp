<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	

	%>
	
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	






<div id="sindhold" style="position:absolute; left:20; top:132; visibility:visible; z-index:100;">
<!-------------------------------Sideindhold------------------------------------->

Kørsels statistik er flyttet til joblog.<br />
Vælg joblog i menuen ovenfor og tilvælg aktivitetstype "kørsel" i søgekriterierne.<br /><br />
Med venlig hilsen<br />
OutZourCE

</div>
<br>
<br>
&nbsp;
<%end if 'validering %>

<!--#include file="../inc/regular/footer_inc.asp"-->