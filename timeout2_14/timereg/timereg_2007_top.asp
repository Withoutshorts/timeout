<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<div id="sindhold" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(1)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:97; visibility:visible;">
	<a href="timereg_akt_2006.asp?showakt=1" class=rmenu target="a"><%=tsa_txt_116 %></a>
	<%if ((level < 8 AND lto <> "bowe") OR (level < 7)) then %>
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="materialer_indtast.asp?id=0&fromsdsk=0&aftid=0" class=rmenu target="_blank"><%=tsa_txt_191 %></a>
	<%end if %>
	
	<!--
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="timereg_akt_2006.asp?showakt=0" class=rmenu target="a"><%=tsa_txt_117 %></a>
	-->
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="joblog_timetotaler.asp" class=rmenu target="_blank"><%=tsa_txt_119 %></a>
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="joblog.asp" class=rmenu target="_blank"><%=tsa_txt_118 %></a>
	
	
	    <%if level = 1 then %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='feriekalender.asp?menu=job' class='rmenu' target="_top">Ferie, Afspad. & Sygdom's kalender</a>
		<%else %>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='feriekalender.asp?menu=job' class='rmenu' target="_top">Ferie & Afspad. kalender</a>
		<%end if %>
	
	
	
	</div>
    
    
    
    

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
