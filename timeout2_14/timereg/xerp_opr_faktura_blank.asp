<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->


<%

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<% 
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	func = request("func")
    thisfile = "erp_opr_faktura_kontojob.asp"
    print = request("print")
    
   %>
	
	
	<div id="sindhold" style="position:absolute; left:0px; top:0px; visibility:visible;">
        <!--<img src="../ill/topshadow3.gif" width="550" height="10" /><br />-->
	<%
	'**********************************************************
	'**************** Blank **************
	'**********************************************************
	 %>
	
	<!--Din faktura oprettelse bliver vist her.-->

	</div>
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->