<!--#include file="../inc/connection/conn_db_inc.asp"-->
<%
	oConn.execute("DELETE FROM Mtid WHERE id= " & request("id") &"")
	
	if request("func") = "afslut" then
	Response.redirect "afslut.asp"
	else
	Response.redirect "ny_indtastning.asp"
	end if
%>
 
   