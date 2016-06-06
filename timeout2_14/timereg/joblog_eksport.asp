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
    
    
    level = session("rettigheder")
    
    hidetimepriser = request.cookies("stat")("hidetimepriser")
    hideenheder = request.cookies("stat")("hideenheder")
    hidegkfakstat = request.cookies("stat")("hidegkfakstat")
    
	'ekspTxt = request("FM_eksportdata")


	'txt1 = request.form("txt1")
	''txt2 = request.form("txt2")
	
	Dim BigTextArea

	For I = 1 To Request.Form("BigTextArea").Count
  		BigTextArea = BigTextArea & Request.Form("BigTextArea")(I)
	Next
	
	'Response.write BigTextArea 
	'Response.Flush
	'Response.end
	
	txt20 = request.form("txt20")
	txt = txt1 & BigTextArea & txt20
	
	ekspTxt = replace(txt, "xx99123sy#z", vbcrlf)
	
	datointerval = request("datointerval")
	
    call TimeOutVersion()
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\joblog_eksport.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				if lto = "bowe" then
				strOskrifter = "Kontakt;Kontakt Id;Job;KA Nr;Böwe kode;Job Nr;Uge;Dato;Fase;Aktivitet;Type;Fakturerbar;Akt. tidslås;Medarbejder;Medarb. Nr;Initialer;Antal;Tid/Klokkeslet;"
				else
				strOskrifter = "Kontakt;Kontakt Id;Job;Job Nr;Uge;Dato;Fase;Aktivitet;Type;Fakturerbar;Medarbejder;Medarb. Nr;Initialer;Antal;Tid/Klokkeslet;"
				end if
				
				if cint(hideenheder) = 0 then
				strOskrifter = strOskrifter &"Enheder;"
				end if
				
				
				if (level <=2 OR level = 6) AND cint(hidetimepriser) = 0 then
				strOskrifter = strOskrifter & "Timepris;Timepris ialt;Valuta;"
				end if 
				
				if level = 1 AND visKost = 1 then 
				strOskrifter = strOskrifter & "Kostpris;Kostpris ialt;Valuta;"
				end if
				
				if lto <> "execon" AND cint(hidegkfakstat) = 0 then
				strOskrifter = strOskrifter  &"Godkendt?;Godkendt af;Faktureret?;"
				end if
				
                
                strOskrifter = strOskrifter  &"Kommentar;"
				
				
				objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				objF.WriteLine(strOskrifter & chr(013))
				objF.WriteLine(ekspTxt)
				
				
				
				
	
	Response.redirect "../inc/log/data/"& file &""	
	%>
	
	
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->

