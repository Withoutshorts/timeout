<%Response.buffer = true%>
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
	
    
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\fakturaer_eksport.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
		
	end if
	
	file = "fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
	
	fakids = request("fakids")
	fakIder = split(fakids, ",") 
	for i = 0 to Ubound(fakIder)
	
		if i = 0 then
		fakidKri = "( f.fid =  " & fakIder(i)
		else
		fakidKri = fakidKri & " OR f.fid =  " & fakIder(i)
		end if
	
	next
	
	fakidKri = fakidKri & ")"
	
	strSQL = "SELECT f.fid, f.faknr, f.dato AS fdato, "_
	&" f.fakdato, f.jobid, f.timer, f.beloeb, "_
	&" f.kommentar, f.faktype, f.parentfak, j.jobnavn, "_
	&" j.jobnr, j.jobknr, Kkundenavn, Kkundenr, tidspunkt AS faktidspkt, f.betalt, f.faktype, "_
	&" betalt, ka.mnavn AS kans, ka.mnr AS kansnr, "_
	&" ja1.mnavn AS jobans1, ja1.mnr AS jans1nr, ja2.mnavn AS jobans2, ja2.mnr AS jans2nr, "_
	& "fd.beskrivelse, fd.antal AS timer, fd.enhedspris AS timepris, fd.aktpris AS belob, "_
	&" ka2.mnavn AS kans2, ka2.mnr AS kansnr2, fd.aktid AS fdaktid, "_
	& "a.navn AS aktnavn, fo.navn AS fomr"_
	&" FROM fakturaer f "_
	&" LEFT JOIN job j ON (j.id = f.jobid)"_ 
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr)"_ 
	&" LEFT JOIN faktura_det fd ON (fd.fakid = f.fid)"_ 
	&" LEFT JOIN medarbejdere ka ON (ka.mid = k.kundeans1)"_ 
	&" LEFT JOIN medarbejdere ka2 ON (ka.mid = k.kundeans2)"_ 
	&" LEFT JOIN medarbejdere ja1 ON (ja1.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere ja2 ON (ja2.mid = j.jobans2)"_
	&" LEFT JOIN aktiviteter a ON (a.id = fd.aktid)"_ 
	&" LEFT JOIN fomr fo ON (fo.id = a.fomr)"_ 
	&" WHERE "& fakidKri &" ORDER BY f.fakdato DESC, tidspunkt DESC"
	
	'Response.write strSQL & "<hr>"
	'Response.flush
	
	strTxtExport = "Aktivitet/Medarbejder;Kontakt;Kontakt id;Kontaktansv.1;Kontaktansv.1 nr.;Kontaktansv.2;Kontaktansv.2 nr.;Jobnavn;Jobnr;Jobans1;Jobans1 nr.;jobans2;Jobans2 nr.;Faktura nr.;Fakdato;Type;Kladde/Godkendt;Aktivitetsnavn;Medarbejdernavn;Medarb. nr.;Timer;Enhedspris;Beløb;Ventetimer;Vist på Faktura?;Forretningsområde;Medarb.linier findes?(sumakt/andre?);Fakturalinie tekst;"
	strTxtExport = strTxtExport & vbcrlf
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 

		
		'*** Export ****
		strTxtExport = strTxtExport & vbcrlf
		strTxtExport = strTxtExport & "A;"
		strTxtExport = strTxtExport & oRec("kkundenavn") &";"
		strTxtExport = strTxtExport & oRec("kkundenr") &";"
		strTxtExport = strTxtExport & oRec("kans") &";"
		strTxtExport = strTxtExport & oRec("kansnr")&";"
		strTxtExport = strTxtExport & oRec("kans2") & ";"
		strTxtExport = strTxtExport & oRec("kansnr2")&";"
		strTxtExport = strTxtExport & oRec("jobnavn") &";"
		strTxtExport = strTxtExport & oRec("jobnr") &";"
		strTxtExport = strTxtExport & oRec("jobans1")&";"
		strTxtExport = strTxtExport & oRec("jans1nr")&";"
		strTxtExport = strTxtExport & oRec("jobans2")&";"
		strTxtExport = strTxtExport & oRec("jans2nr")&";"
		strTxtExport = strTxtExport & oRec("faknr") &";"
		strTxtExport = strTxtExport & formatdatetime(oRec("fakdato")) &";"
		
		faktypeThis = "-"
		select case oRec("faktype")
		case 1
		faktypeThis = "Kreditnota"
		case 2
		faktypeThis = "Rykker"
		case else
		faktypeThis = "Faktura"
		end select  
		
		
		strTxtExport = strTxtExport & faktypeThis &";"
		strTxtExport = strTxtExport & oRec("betalt") &";"
		strTxtExport = strTxtExport & oRec("aktnavn") &";"
		
		strTxtExport = strTxtExport & "-;"
		strTxtExport = strTxtExport & "-;"
		strTxtExport = strTxtExport & oRec("timer") &";"
		strTxtExport = strTxtExport & oRec("timepris") &";"
		strTxtExport = strTxtExport & oRec("belob") &";"
		strTxtExport = strTxtExport & ";"
		strTxtExport = strTxtExport & "1;"
		strTxtExport = strTxtExport & oRec("fomr") &";"
		
		
		
		
				if len(oRec("timepris")) <> 0 then
				tprisThis = replace(oRec("timepris"), ",",".")
				else
				tprisThis = 0
				end if
				
				if len(oRec("fdaktid")) <> 0 then
				aktidThis = oRec("fdaktid")
				else
				aktidThis = 0
				end if
				
				if len(oRec("fid")) <> 0 then
				fakturaidThis = oRec("fid")
				else
				fakturaidThis = 0
				end if
				
				'** Medarbejdere ***
				strSQL2 = "SELECT fms.fak AS fmstimer, fms.showonfak AS fmsshowonfak,"_
				&" fms.enhedspris AS fmsenhedspris, fms.beloeb AS fmsbelob, fms.tekst, "_
				&" m.mnavn, m.mnr, fms.venter AS fmsventer "_
				&" FROM fak_med_spec fms "_
				&" LEFT JOIN medarbejdere m ON (m.mid = fms.mid) "_
				&" WHERE fms.fakid = "& fakturaidThis &" AND fms.aktid = "& aktidThis &" AND fms.enhedspris = " & tprisThis
				
				'Response.write strSQL2
				'Response.flush
				
				oRec2.open strSQL2, oConn, 3 
				m = 0
				while not oRec2.EOF 
					
					if m = 0 then
					strTxtExport = strTxtExport & "1;"
					strTxtExport = strTxtExport & oRec("beskrivelse") &";"
					strTxtExport = strTxtExport & vbcrlf
					end if
					
					strTxtExport = strTxtExport & "M;"
					strTxtExport = strTxtExport & oRec("kkundenavn") &";"
					strTxtExport = strTxtExport & oRec("kkundenr") &";"
					strTxtExport = strTxtExport & oRec("kans") &";"
					strTxtExport = strTxtExport & oRec("kansnr")&";"
					strTxtExport = strTxtExport & oRec("kans2") & ";"
					strTxtExport = strTxtExport & oRec("kansnr2")&";"
					strTxtExport = strTxtExport & oRec("jobnavn") &";"
					strTxtExport = strTxtExport & oRec("jobnr") &";"
					strTxtExport = strTxtExport & oRec("jobans1")&";"
					strTxtExport = strTxtExport & oRec("jans1nr")&";"
					strTxtExport = strTxtExport & oRec("jobans2")&";"
					strTxtExport = strTxtExport & oRec("jans2nr")&";"
					strTxtExport = strTxtExport & oRec("faknr") &";"
					strTxtExport = strTxtExport & formatdatetime(oRec("fakdato")) &";"
					strTxtExport = strTxtExport & faktypeThis &";"
					strTxtExport = strTxtExport & oRec("betalt") &";"
					strTxtExport = strTxtExport & oRec("aktnavn") &";"
					
					strTxtExport = strTxtExport & oRec2("mnavn") &";"
					strTxtExport = strTxtExport & oRec2("mnr")&";"
					strTxtExport = strTxtExport & oRec2("fmstimer") &";"
					strTxtExport = strTxtExport & oRec2("fmsenhedspris") &";"
					strTxtExport = strTxtExport & oRec2("fmsbelob") &";"
					strTxtExport = strTxtExport & oRec2("fmsventer") &";"
					strTxtExport = strTxtExport & oRec2("fmsshowonfak") &";"
					strTxtExport = strTxtExport & oRec("fomr") &";"
					strTxtExport = strTxtExport & "-;"
					strTxtExport = strTxtExport & oRec2("tekst") &";"
					strTxtExport = strTxtExport & vbcrlf
				
				m = m + 1	
				oRec2.movenext
				wend
				oRec2.close 
		
				if m = 0 then
				strTxtExport = strTxtExport & "0;"
				strTxtExport = strTxtExport & oRec("beskrivelse") & ";"
				strTxtExport = strTxtExport & vbcrlf
				end if
		
	
	oRec.movenext
	wend
	oRec.close
	
	
	'strTxtExport = Replace(strTxtExport, "ø", "&oslash;")
	'strTxtExport = Replace(strTxtExport, "Ø", "&Oslash;")
	'strTxtExport = Replace(strTxtExport, "æ", "&aelig;")
	'strTxtExport = Replace(strTxtExport, "Æ", "&AElig;")
	'strTxtExport = Replace(strTxtExport, "å", "&aring;")
	'strTxtExport = Replace(strTxtExport, "Å", "&Aring;")
	
	
	objF.WriteLine(strTxtExport)
	
	
	objF.close

	Response.redirect "../inc/log/data/"& file &""	
	
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
		
