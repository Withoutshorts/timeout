	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<%
	'Response.contenttype = "application/octet-stream"
	
	select case request("action")
	case "1"
	medArbId = Request("mid")
	areJobPrinted = "0"
		f = 0
		Redim brugergruppeId(f)
		
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& medArbID &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 3
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = oRec("Mid")
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close

	for intcounter = 0 to f - 1  
  	sortby = "jobnr DESC"
	
	strSQL = "SELECT job.id, jobnr, jobnavn, jobknr, kkundenavn, aktiviteter.id AS Aid, aktiviteter.navn, aktiviteter.job  FROM job, aktiviteter, kunder WHERE jobstatus = 1 AND aktiviteter.job = job.id AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe1 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id, jobnr, jobnavn, jobknr,  kkundenavn, aktiviteter.id AS Aid, aktiviteter.navn, aktiviteter.job  FROM job, aktiviteter, kunder WHERE jobstatus = 1 AND aktiviteter.job = job.id AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe2 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id, jobnr, jobnavn, jobknr,  kkundenavn, aktiviteter.id AS Aid, aktiviteter.navn, aktiviteter.job  FROM job, aktiviteter, kunder WHERE jobstatus = 1 AND aktiviteter.job = job.id AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe3 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id, jobnr, jobnavn, jobknr,  kkundenavn, aktiviteter.id AS Aid, aktiviteter.navn, aktiviteter.job  FROM job, aktiviteter, kunder WHERE jobstatus = 1 AND aktiviteter.job = job.id AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe4 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id, jobnr, jobnavn, jobknr,  kkundenavn, aktiviteter.id AS Aid, aktiviteter.navn, aktiviteter.job  FROM job, aktiviteter, kunder WHERE jobstatus = 1 AND aktiviteter.job = job.id AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe5 = "&brugergruppeId(intcounter)&""_
	&" ORDER BY "& sortby &""
	
	oRec.Open strSQL, oConn, 3  
	
	While Not oRec.EOF
	areJobPrinted = instr(jobPrinted, oRec("Jobnr"))
	 
		if areJobPrinted = "0" then
		jobPrinted = jobPrinted &", "& oRec("Jobnr")  
		Response.write oRec("id")& ", "& oRec("jobnavn") & chr(13) & chr(10)
		end if
	
	oRec.movenext
	wend
	
	oRec.close
	
	next
	case "2"
	varjobnr = request("job_id")
	medArbId = 1
	
	f = 0
		Redim brugergruppeId(f)
		
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& medArbID &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 3
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = oRec("Mid")
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close

	for intcounter = 0 to f - 1  
	
	strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, id AS akt, job, navn FROM aktiviteter WHERE job = "& varjobnr &" AND projektgruppe1 = "& brugergruppeId(intcounter) &" OR job = "& varjobnr &" AND projektgruppe2 = "& brugergruppeId(intcounter) &" OR job = "& varjobnr &" AND projektgruppe3 = "& brugergruppeId(intcounter) &" OR job = "& varjobnr &" AND projektgruppe4 = "& brugergruppeId(intcounter) &" OR job = "& varjobnr &" AND projektgruppe5 = "& brugergruppeId(intcounter) &""
	oRec3.open strSQL3, oConn, 3
	
	While not oRec3.EOF 
		
		Response.write oRec3("akt")& ", "& oRec3("navn") & chr(13) & chr(10)
	
	oRec3.MoveNext
	Wend 
	
	oRec3.Close
	
	next
	
	
case "3"
varjobId = request("job_id")
varAktId = request("ac_id")
varDato = request("dato")
strMnr = 1
varTimer = request("timer")
'datoUS = cdate(convertDateUS(varDato))

	
	strSQL3 = "SELECT job, navn FROM aktiviteter WHERE job = "& varjobId &"" 
	oRec3.open strSQL3, oConn, 3
	
	if not oRec3.EOF then
		strAktNavn = oRec3("navn")
	end if
	
	oRec3.Close
	
	strSQL3 = "SELECT jobnavn, jobnr,  jobknr, Kkundenavn, Kkundenr, Kid FROM job, kunder WHERE id = "& varjobId &" AND kunder.Kid = jobknr" 
	oRec3.open strSQL3, oConn, 3
	
	if not oRec3.EOF then
		strJobnr = oRec3("jobnr")
		strJobnavn = oRec3("jobnavn")
		strKundeNavn = oRec3("Kkundenavn")
		intKkundeNr = oRec3("Kid")
	end if
	
	oRec3.Close

	oConn.Execute("INSERT INTO timer (TJobnr, TJobnavn, Tmnr, Tmnavn, Tdato, Timer, Tknavn, Tknr, TAktivitetId, TAktivitetNavn, Tfaktim, Taar, TimePris, TasteDato) VALUES"_
	& "(" & strJobnr & ", '"&strJobnavn&"', " & strMnr & ", 'Søren Karlsen', #5/3/2002#, "& varTimer &", '"&strKundeNavn&"', "& intKkundeNr &", "& varAktId &", '"& strAktNavn &"', 1, 2002, 650, '"& month(now)&"/"&day(now)&"/"&year(now)&"')")
	
	
	'strSQL3 = "SELECT * FROM timer" 
	'oRec3.open strSQL3, oConn, 3
	
	'oRec3.movelast
	'if not oRec3.EOF then
		'Response.write "jobnavn: " & oRec3("TJobnavn") & "<br>"
		'Response.write "aktivitet: " & oRec3("TAktivitetNavn") & "<br>"
		'Response.write "dato: " & oRec3("Tdato") & "<br>"
	'end if
	
	'oRec3.Close

end select
%>


 




