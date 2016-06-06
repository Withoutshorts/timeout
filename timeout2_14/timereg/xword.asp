<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	if len( request("FM_start_dag")) <> 0 then
	strMrd = request("FM_start_mrd")
	strDag = request("FM_start_dag")
	strAar = right(request("FM_start_aar"),2) 
	strDag_slut = request("FM_slut_dag")
	strMrd_slut = request("FM_slut_mrd")
	strAar_slut = right(request("FM_slut_aar"),2)
	else
	strMrd_temp = dateadd("m", -1, date)
	strMrd = month(strMrd_temp)
	strDag_temp = dateadd("d", -30, date)
	strDag = day(strDag_temp)
	if strMrd <> 12 then
	strAar= right(year(now), 2)
	else
	strAar_temp = dateadd("yyyy", -1, date)
	strAar = right(year(strAar_temp),2)
	end if
	
	strDag_slut = day(now)
	strMrd_slut = month(now)
	strAar_slut = right(year(now), 2)
	end if
	
	
	'** Indsætter cookie **
	Response.Cookies("datoer")("st_dag") = strDag
	Response.Cookies("datoer")("st_md") = strMrd
	Response.Cookies("datoer")("st_aar") = strAar
	Response.Cookies("datoer")("sl_dag") = strDag_slut
	Response.Cookies("datoer")("sl_md") = strMrd_slut
	Response.Cookies("datoer")("sl_aar") = strAar_slut
	Response.Cookies("datoer").Expires = date + 10		
	
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	
	'weekselector = request("weekselector")
	'strWSelStartDato = request("strWSelStartDato")
	'strWSelEndDato = request("strWSelEndDato")
	linket = request("linket")
	FM_medarb = request("FM_medarb")
	
	selmedarb = FM_medarb 'request("selmedarb")
	if len(selmedarb) = "0" then
	selmedarb = 0
	else
	selmedarb = selmedarb
	end if
	
	selaktid = request("selaktid")
	FM_job = request("FM_job")
	intJobnr = FM_job 'request("jobnr")
	
	thisfile = "word"
	
	'*** Finder de medarbejdere og job der er valgt ***
		Dim intMedArbVal 
		Dim b
		Dim intJobnrKriValues 
		Dim i
				
					intMedArbVal = Split(selmedarb, ", ")
					For b = 0 to Ubound(intMedArbVal)
						intJobnrKriValues = Split(intJobnr, ", ")
						For i = 0 to Ubound(intJobnrKriValues)
						if selmedarb <> "0" then
						jobMedarbKri = jobMedarbKri & " Tmnr = " & intMedArbVal(b) &" AND Tjobnr = " & intJobnrKriValues(i) &" AND Tfaktim <> 5 OR "
						else
						jobMedarbKri = jobMedarbKri & " Tmnr <> 0 AND Tjobnr = " & intJobnrKriValues(i) &" AND Tfaktim <> 5 OR "
						end if
						next
					next
					intJobMedarbKriCount = len(jobMedarbKri)
					trimIntJobMedarbKri = left(jobMedarbKri, (intJobMedarbKriCount-3))
					jobMedarbKri = trimIntJobMedarbKri & " AND "
		
		
		if len(request("selaktid")) <> "0" then
		selaktidKri = "AND TAktivitetId = " & selaktid &""
		else
		selaktidKri = ""
		end if

	
				
				filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
				filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
				
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\word.asp" then
										
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\statexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\statexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				
				else
					
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\statexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\statexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
					
				end if
				
				file = "statexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
	
	
	
				
				
			'	Set objFSO = server.createobject("Scripting.FileSystemObject")
				
			'	Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\joblog_"&lto&".txt", True, False)
			'	'Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblog_"&lto&".txt", True, False)
				
			'	Set objNewFile = nothing
			'	Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\joblog_"&lto&".txt", 8)
			'	'Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblog_"&lto&".txt", 8)
				
			'	file = "joblog_"&lto&".txt"
				
				'*** 1 linie ****
				objF.WriteLine("Uge;Dato;Kunde;Jobnr;Jobnavn;Aktivitet;Medarbejder;Timer / Enheder;Heraf fakturerbare timer;Reg. Timepris (uanset fastpris/lbn.tim.);Gns.fak. Timepris;Kommentar;")
		
	
	
	'**** Main SQL ***
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom FROM timer WHERE ("& jobMedarbKri &" Tid > 0 "& selaktidKri &") AND tdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"' ORDER BY Tdato DESC, Tmnavn"
	oRec.open strSQL, oConn, 0, 1
	
	v = 0
	while not oRec.EOF
			
			strWeekNum =  datepart("ww", oRec("Tdato"))
			id = 1
			
					
					
					if len(oRec("Timer")) <> 0 then
					timerthis = oRec("Timer") 'replace(oRec("Timer"), ",",".")
					else
					timerthis = 0
					end if
					
					if oRec("Tfaktim") = 1 then
					timerFakbare = timerthis
					else
					timerFakbare = 0
					end if
					
						
						
						
						strSQL2 = "SELECT avg(fm.enhedspris) AS medfakbel FROM job, fakturaer f, fak_med_spec fm WHERE "_
						&"(job.jobnr = "& oRec("tjobnr") &") AND (fakdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"' AND jobid = job.id) AND (fm.fakid = f.fid AND fm.mid = "& oRec("Tmnr") &" AND aktid = "& oRec("TAktivitetId") &")"
						
						'if v = 0 then
						'Response.write strSQL2
						'end if
						faktimepris = 0
						
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						faktimepris = oRec2("medfakbel")
						end if
						oRec2.close 
						
						
						
					objF.WriteLine(strWeekNum&";"&convertDate(oRec("Tdato"))&";"&oRec("Tknavn")&";"&oRec("Tjobnr")&";"&oRec("Tjobnavn")&";"&oRec("Anavn")&";"&oRec("Tmnavn")&";"&timerthis &";"&timerFakbare &";"&oRec("TimePris") &";"& faktimepris &";"""&oRec("Timerkom")&""";")
					totTimer = totTimer + oRec("timer")
					
					if oRec("Tfaktim") = 1 then
					totOms = totOms + (oRec("timer") * oRec("TimePris"))
					else
					totOms = totOms
					end if
					
					v = v + 1
					
					
			
	oRec.movenext
	wend	  
	oRec.Close
	Set oRec = Nothing
	

objF.writeLine("Periode: "& formatdatetime(strDag&"/"&strMrd&"/"&strAar, 0) &" til  "& formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2))
objF.WriteLine("Timer total: " & totTimer &" Omsætning (kun fakbare timer medregnet):" & formatcurrency(totOms, 2))

objF.close

Response.redirect "../inc/log/data/"& file &""
%>
<!--
<br>
<br>
 <a href="c:\joblog_<=intJobnr%>.txt">Åben dokument</a>-->
<%	
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->

