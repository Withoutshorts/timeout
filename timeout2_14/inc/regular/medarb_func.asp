<%

	
	
	function medarbtopmenu()
		%>
		<br><a href='medarb.asp' class='rmenu'>Medarbejdere</a>
		<%
		if level = 1 then 
		%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='bgrupper.asp' class='rmenu'>Brugergrupper</a>
        <%end if%>

        <%
		if (level = 1 AND lto <> "mi") OR (level = 1 AND (session("mid") = 2 OR session("mid") = 29 OR session("mid") = 1 OR session("mid") = 38))  then 
		%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href='medarbtyper.asp' class='rmenu'>Medarbejdertyper</a>
		<%end if%>
		<br>&nbsp;
		<%
	end function
	
%>