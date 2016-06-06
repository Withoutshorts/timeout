
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
ekspTxt = request("FM_eksportdata")
	
				'** Opdater Eksport Status **** 
				aftaler = split(request("FM_aftaler"), ",")
				for i = 0 to Ubound(aftaler)
				oConn.execute("UPDATE serviceaft SET erfaktureret = 1 WHERE id = "& aftaler(i) &"")
				next	
				
                call TimeOutVersion()
				
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				'Response.write request.servervariables("PATH_TRANSLATED")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\fak_serviceaft_eksport.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\aftaler_"&lto&".txt", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\aftaler_"&lto&".txt", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\aftaler_"&lto&".txt", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\aftaler_"&lto&".txt", 8)
				end if
				
				
				
				file = "aftaler_"&lto&".txt"
				objF.WriteLine(ekspTxt)
				
	
	
	%>
	<div id="sindhold" style="position:absolute; left:20; top:20; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
    <td valign="top"><a href="../inc/log/data/<%=file%>" class=vmenu>Data er nu eksporteret til denne fil.&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br>
	<br>
	Gem filen ved at klikke på filmenuen og herefter "gem som".<br>
	Åbn det program du ønsker at hente txt filen ind i og importer.<br><br>
	
	</td>
	</tr></table>
	</div>
	
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->