
<%


	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>
	
<!-- frames -->
	
	<frameset rows="126,*" framespacing="0" border="0" name="main2">
		
		<frame src="timereg_2007_top.asp" name="top" id="top" frameborder="0" scrolling="auto" marginwidth="0" marginheight="0">
	    
	    
	<frameset  cols="230,*" framespacing="0" name="main" id="main">
		
		<frame src="timereg_2006.asp" name="t" id="t" frameborder="0" scrolling="auto" marginwidth="0" marginheight="0">
	    <frame name="a" id="a" src="timereg_akt_2006.asp" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0">
		
	    
	</frameset>
	
	</frameset>
	
	
	<%end if %>






