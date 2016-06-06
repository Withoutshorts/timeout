<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->

<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	menu = "erp"
	%>
	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<div id="sindhold" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call erpmainmenu(1)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call erptopmenu()
	%>
	</div>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
