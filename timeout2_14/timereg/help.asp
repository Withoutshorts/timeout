<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<body topmargin=0 leftmargin=0>



    <%

    call menu_2014()
    
    'oimg = "lifebelt.png"
	oleft = 30
	otop = 82
	owdt = 500
	oskrift = "Hjælp og FAQ til TimeOut " & year(now)
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
     %>
	
    <br /><br />
    <br /><br /><br />
	
	<%call helpandguides(lto, 1)  %>
	
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
&nbsp;
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->

