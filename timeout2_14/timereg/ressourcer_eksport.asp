<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<script>
	function changeZoom2(oSel) {
  	newZoom= oSel.options[oSel.selectedIndex].innerText
	printarea.style.zoom=newZoom+'%';
	//oCode.innerText='zoom: '+newZoom+'';	
  	} 
</SCRIPT>

<!--#include file="../inc/regular/header_hvd_inc.asp"-->

<%	
	txt1 = request.form("txt1")
	'txt2 = request.form("txt2")
	
	Dim BigTextArea

	For I = 1 To Request.Form("BigTextArea").Count
  		BigTextArea = BigTextArea & Request.Form("BigTextArea")(I)
	Next
	
	txt20 = request.form("txt20")
	csvTxt_samlet = txt1 & BigTextArea & txt20
	
	'datointerval = request("datointerval")

    call TimeOutVersion()
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\ressourcer_eksport.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
		
	end if
	
	file = "ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
	
	
	objF.WriteLine(csvTxt_samlet)
	
	
	objF.close

	Response.redirect "../inc/log/data/"& file &""	

%>	
<!--#include file="../inc/regular/footer_inc.asp"-->