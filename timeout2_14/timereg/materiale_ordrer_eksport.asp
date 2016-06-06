
<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->


<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
%>
<!--#include file="../inc/regular/header_hvd_inc.asp"-->
<%


	
	
	Dim BigTextArea

	For I = 1 To Request.Form("BigTextArea").Count
  		BigTextArea = BigTextArea & Request.Form("BigTextArea")(I)
	Next
	
	txt20 = request.form("txt20")
	txt = txt1 & BigTextArea & txt20
	
	ekspTxt = replace(txt, "xx99123sy#z", vbcrlf)
	
	
    call TimeOutVersion()
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\materiale_ordrer_eksport.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\materiale_ordrer_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\materiale_ordrer_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\materiale_ordrer_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\materiale_ordrer_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "materiale_ordrer_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				objF.WriteLine(ekspTxt)
				
	
	Response.redirect "../inc/log/data/"& file &""	
	%>
	
	
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->