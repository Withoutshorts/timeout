<!--#include file="../inc/regular/global_func.asp"-->
<%
ekspTxt = request("FM_eksportdata")
	
				
				
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
				
				'objF.writeLine("Periode afgrænsning: "& formatdatetime(strDag&"/"&strMrd&"/"&strAar, 0) &" til "& formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 0))
				'objF.WriteLine("::KUNDENR")
				'objF.WriteLine(";;VARENR;ANTAL;PRIS;RABAT*;AFTALENAVN;PERIODE;NOTAT")
				objF.WriteLine(ekspTxt)
				objF.WriteLine("==")
	%>
	<a href="../inc/log/data/<%=file%>" class=vmenu>Data er nu eksporteret til denne fil.&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a><br>
	<br>
	Gem filen ved at klikke på filmenuen og herefter "gem som".<br>
	Åbn det program du ønsker at hente txt filen ind i og importer.<br><br>
