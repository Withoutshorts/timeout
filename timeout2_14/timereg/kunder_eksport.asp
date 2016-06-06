<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/convertDate.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	if len(request("kpers")) <> 0 then
	kpers = request("kpers")
	else
	kpers = 0
	end if

     call TimeOutVersion()
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\kunder_eksport.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\kundeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\kundeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\kundeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\kundeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
		
	end if
	
	file = "kundeexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
	
	
	'Set objFSO = server.createobject("Scripting.FileSystemObject")
				
	'Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\kundeexp_"&lto&".txt", True, false)
	''Set objNewFile = objFSO.createTextFile("e:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\kundeexp_"&lto&".txt", True, False)
	
	'Set objNewFile = nothing
	'Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\kundeexp_"&lto&".txt", 8, -1)
	''Set objF = objFSO.OpenTextFile("e:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\kundeexp_"&lto&".txt", 8)
	
	'file = "kundeexp_"&lto&".txt"
	
	
	    'if cint(kpers) = -1 then
		
		strTxtExport =  "Linjeoverskrift;Kontakt;Kontakt id;Kontaktperson/filial;Adresse; "_
		& "Postnr;By;Land;Telefon;Mobil;Fax;"_
		& "Email;Titel;Segment;Afdeling;Rabat;EAN;Interesse;Beskrivelse;Konto;CVR;Bank;Iban;Swift;NCA kode;==" & vbcrlf
		
		'else
		
		'strTxtExport =  "Linjeoverskrift;Kontakt;Kontakt id;Kontaktperson;Adresse; "_
		'& "Postnr;By;Land;Telefon;Telefon Dir.;Mobil;Fax;"_
		'& "Email;Titel;Segment;Afdeling;==" & vbcrlf
		
		
		'end if
	
    dim strEan
    strEan = ""
	
	kids = request("kids")
	kunderIder = split(kids, ",") 
	for i = 0 to Ubound(kunderIder)
	
	strSQL = "SELECT Kid, Kkundenavn, Kkundenr, Kdato, "_
	& "adresse, postnr, city, land, telefon, "_
	& "telefonmobil, telefonalt, fax, email, ean, url, "_
	& "ketype, beskrivelse, "_
	& "regnr, kontonr, cvr, bank, swift, iban, useasfak, hot, "_
	& "logo, ktype, kundeans1, kundeans2, kt.navn AS kundetypenavn, kt.rabat, nace "_
	&" FROM kunder LEFT JOIN kundetyper kt ON (kt.id = ktype) WHERE Kid=" & kunderIder(i)
	
	'Response.write strSQL & "<hr>"
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	intKid = oRec("Kid") 
	strKundenavn = oRec("Kkundenavn")
	strKnr = oRec("Kkundenr")
	strDato = oRec("Kdato")
	strLastUptDato = oRec("Kdato")
	'strEditor = oRec("editor")
	strAdr = oRec("adresse")
	strPostnr = oRec("postnr")
	strCity = oRec("city")
	strLand = oRec("land")
	
    

    if len(trim(oRec("telefon"))) <> 0 then
    strTlf = cstr(oRec("telefon"))
    else
    strTlf = ""
    end if

	strMobil = oRec("telefonmobil")
	strAlttlf = oRec("telefonalt")
	
    strFax = oRec("fax")
	strEmail = oRec("email")
    
    
    if len(trim(oRec("ean"))) <> 0 then
    strEan = cstr("EAN "&oRec("ean"))
    else
    strEan = ""
    end if

	'strWWW = oRec("url")
	ketype = oRec("ketype")
	'strKomm = oRec("beskrivelse")
	intregnr = oRec("regnr")
	intkontonr = oRec("kontonr")
	intCVR = oRec("cvr")
	strBank = oRec("bank")
	strIban = oRec("iban")
	strSwift = oRec("swift")
	logo = oRec("logo")
	intKtype = oRec("ktype")
	strKundeType = oRec("kundetypenavn")
	strNACE = oRec("nace")
	rabat = oRec("rabat")
	hot = oRec("hot")
	
    strBesk = oRec("url")

		'*** Export ****
		'objF.WriteLine("Id "&  chr(009) &" Kundenavn "&  chr(009) &"Kundenr "& chr(009) &"Jobnavn "& chr(009)&"Aktivitet "& chr(009) &"Kunde "& chr(009) &"Timer / Enheder"& chr(009) &"Heraf fakturerbare timer"& chr(009) &"Reg. Timepris (uanset fastpris/lbn.tim.)"& chr(009) &"Gns.fak. Timepris" & chr(009) & "Medarbejder "& chr(009) &"Kommentar ")
        
        
        
		if len(trim(strAdr)) <> 0 then
		call htmlparseCSV(strAdr)
		strAdr = replace(htmlparseCSVtxt, "&vbcrlf", "")
        'strAdr = replace(strAdr, "&vbcrlf", "")
        else
        strAdr = ""
        end if

        if len(trim(strBesk)) <> 0 then
		call htmlparseCSV(strBesk)
		 strBesk = replace(htmlparseCSVtxt, "&vbcrlf", "")
        else
         strBesk = ""
        end if

       
		
        'Response.flush

		'if cint(kpers) = -1 then
		
		strTxtExport = strTxtExport & "Kontakt;"& Chr(34) & trim(strKundenavn) & Chr(34) &";"& Chr(34) & strKnr & Chr(34) &";;"& Chr(34) & trim(strAdr) & Chr(34) &";"_
		& Chr(34) & strPostnr & Chr(34) &";"& Chr(34) & strCity & Chr(34) &";"& Chr(34) & strLand & Chr(34) &";"& Chr(34) & strTlf & Chr(34) &";"& Chr(34) & strMobil & Chr(34) &";"& Chr(34) & strFax & Chr(34) &";"_
		& Chr(34) & strEmail & Chr(34) &";;"& Chr(34) & strKundeType & Chr(34) &";;"& Chr(34) & rabat & Chr(34) &";"& Chr(34) & replace(strEan, "EAN ", "") & Chr(34) &";"& Chr(34) & hot & Chr(34) &";"& Chr(34) & strBesk & Chr(34) &";"& Chr(34) & intregnr & intkontonr & Chr(34) &";"_
        & Chr(34) & intCVR & Chr(34) &";"& Chr(34) & strBank & Chr(34) &";"_
		& Chr(34) & strIban & Chr(34) &";" & Chr(34) & strSwift & Chr(34) & ";"& Chr(34) & strNACE & Chr(34) &";"& vbcrlf 
		
		'end if
		
		'strTxtExport = strTxtExport & "Id, Navn, Kunde, "_
		'&" Titel, Afdeling, Adresse, Postnr, By, Land, "_
		'&" Telefon dir., Mobil, Email" & vbcrlf
		
		
		if cint(kpers) = 1 then
		'*** Kontaktpersoner ***'
		strSQLKpers = "SELECT p.id, p.kundeid, p.navn, p.adresse, p.postnr, p.town, "_
		&" p.land, p.afdeling, p.email, p.password, p.dirtlf, p.mobiltlf, p.privattlf, "_
		&" p.beskrivelse, p.titel FROM kontaktpers p "_
		&" WHERE p.kundeid = "& intKid &" AND p.navn <> '0' ORDER BY p.navn"  '"& useThisKpers
		
		t = 0
		'Response.write strSQLKpers &"<br><br>"
		'Response.flush
		
		oRec2.open strSQLKpers, oConn, 3	
		while not oRec2.EOF 			
		
		'intKpersId = oRec2("id")
		strKpers = oRec2("navn") 'strKpers1
		strKperstit = oRec2("titel") 'str_titel_kpers1
		'strKpersTlf = oRec2("privattlf")
		strKpersmTlf = oRec2("mobiltlf")
		strKpersdTlf = oRec2("dirtlf")
		strKpersEmail = oRec2("email")
		'strKpersEan = oRec2("ean")
		'strKperspw = oRec2("password")
		
        

        if len(trim(oRec2("adresse"))) <> 0 then
        
        strAdr_kpers = oRec2("adresse")
		strPostnr_kpers = oRec2("postnr")
		strCity_kpers = oRec2("town")
		strLand_kpers = oRec2("land")

        else

        strAdr_kpers = strAdr
		strPostnr_kpers = strPostnr
		strCity_kpers = strCity
		strLand_kpers = strLand
        
      
        end if
		
        strAf_kpers = oRec2("afdeling")
		'strBeskkpers = oRec2("beskrivelse")
		
		
		strTxtExport = strTxtExport &"Kontaktpers./filial;"& Chr(34) & trim(strKundenavn) & Chr(34) &";"& Chr(34) & strKnr & Chr(34) &";" & Chr(34) & strKpers & Chr(34) &";"_
		& Chr(34) & trim(strAdr_kpers) & Chr(34) & ";" & Chr(34) & strPostnr_kpers & Chr(34) & ";"_
		& Chr(34) & strCity_kpers & Chr(34) &";" & Chr(34) & strLand_kpers & Chr(34) & ";" & Chr(34) & strKpersdTlf & Chr(34) &";" & Chr(34) & strKpersmTlf & Chr(34) & ";;"_
		& Chr(34) & strKpersEmail & Chr(34) & ";"& Chr(34) & strKperstit & Chr(34) &";"& Chr(34) & strKundeType & Chr(34) &";"& Chr(34) & strAf_kpers & Chr(34) &";" & vbcrlf
		
		oRec2.movenext
		wend
		oRec2.close
		
		
		'strTxtExport = strTxtExport & vbcrlf & vbcrlf & vbcrlf
		
		end if 'kpers
		
		
		
	end if
	oRec.close
	
	next
	
	'strTxtExport = Replace(strTxtExport, "ø", "&oslash;")
	'strTxtExport = Replace(strTxtExport, "Ø", "&Oslash;")
	'strTxtExport = Replace(strTxtExport, "æ", "&aelig;")
	'strTxtExport = Replace(strTxtExport, "Æ", "&AElig;")
	'strTxtExport = Replace(strTxtExport, "å", "&aring;")
	'strTxtExport = Replace(strTxtExport, "Å", "&Aring;")
	
	objF.WriteLine(strTxtExport)

	'Response.write strTxtExport
	'Response.flush
	objF.close


    'Response.end
	Response.redirect "../inc/log/data/"& file &""	
	
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
		
