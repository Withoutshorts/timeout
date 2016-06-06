<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else


    select case lto 
    case "epi_uk"
         Session.LCID = 1030
    end select
    
	
	
	if len(request("visning")) <> 0 then
	visning = request("visning")
	else
	visning = 0
	end if
    
    
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
    sub joblogIFog_udlag


        

                '** Materialer ***
				strSQLmat = "SELECT fmat.matfakid, fmat.matid, fmat.matantal, fmat.matnavn, fmat.matvarenr, matfrb_mid, matfrb_id, "_
				&" fmat.matenhed, fmat.matenhedspris, fmat.matrabat, fmat.matbeloeb, fmat.matshowonfak, m.mnavn, m.mnr, m.init, mf.forbrugsdato,"_
				&" fmat.valuta, fmat.kurs, v4.valutakode AS v4valutakode "_
				&" FROM fak_mat_spec fmat "_
                &" LEFT JOIN medarbejdere m ON (m.mid = matfrb_mid) "_
                &" LEFT JOIN materiale_forbrug mf ON (mf.id = matfrb_id) "_
				&" LEFT JOIN valutaer v4 ON (v4.id = fmat.valuta)"_
	            &" WHERE "_
				&" matfakid = "& lastFakID & " ORDER BY mf.forbrugsdato"
				
				'Response.write strSQLmat
				'Response.flush
				
				oRec2.open strSQLmat, oConn, 3 
				mfu = 0
				while not oRec2.EOF 
					
					'if mfu = 0 then
					'strTxtExport = strTxtExport & vbcrlf
					'end if

                    matdatoAAR = datepart("yyyy", oRec2("forbrugsdato"), 2,2) 
                    matdatoMD = datepart("m", oRec2("forbrugsdato"), 2,2) 
					
					strTxtExport = strTxtExport & vbcrlf &  "Materialer/Udlæg;;"
					strTxtExport = strTxtExport & replace(lastFakStamoplys, "#", oRec2("forbrugsdato") &";"& matdatoMD &";"& matdatoAAR)
					
                    mat_enhprisSigned = lastFaksign&((oRec2("matenhedspris"))/1)
                    mat_antalSigned = lastFaksign&((oRec2("matantal"))/1)
                    mat_belobSigned = lastFaksign&((oRec2("matbeloeb"))/1)

					'** Kurs på Faktura SKAL benyttes til sammenligningen, **'
					'** da sumbeløbet på medabl linie er beregnet udfra **'
					'** kurs på faktura ***'

                    if len(trim(mat_belobSigned)) <> 0 then
                    mat_belobSigned = mat_belobSigned
                    else
                    mat_belobSigned = 0
                    end if

					call beregnValuta(mat_belobSigned,lastFakKurs,100)
                    
					

					strTxtExport = strTxtExport & oRec2("matnavn") &";"
                    strTxtExport = strTxtExport & oRec2("mnavn") &";" & oRec2("mnr") &";" & oRec2("init") &";"
					'strTxtExport = strTxtExport & oRec2("matvarenr")&";;"
					strTxtExport = strTxtExport & mat_antalSigned &";"
					strTxtExport = strTxtExport & mat_enhprisSigned &";"
					strTxtExport = strTxtExport & oRec2("v4valutakode") &";"
					
					
					strTxtExport = strTxtExport & lcase(oRec2("matenhed")) &";"
					strTxtExport = strTxtExport & oRec2("matrabat") &";"
					strTxtExport = strTxtExport & formatnumber(valBelobBeregnet/1,2) &";"& basisValISO &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport & oRec2("matshowonfak") &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";;"
					'strTxtExport = strTxtExport & vbcrlf
				
				mfu = mfu + 1	
				oRec2.movenext
				wend
				oRec2.close 

                if lastFakType = "job" then

                if len(trim(LastJobID)) <> 0 then
                LastJobID = LastJobID
                else
                LastJobID = 0
                end if

                '** Kun job fakturaer ***'
                '** Joblog i fakturaperiode på aktiviteter der ikke er med på faktura ***
                '** Alle timer der regnes med i daglig timereg. ***'
				strSQLjoblogIF = "SELECT timer AS sumjoglogakttimer, tmnr, tjobnr, taktivitetid, timepris,"_
                &" a.navn AS aktnavn, m.mnavn, m.mnr, m.init, t.valuta, t.kurs, tdato, aty.aty_enh "_
                &" FROM timer t "_
                &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) "_
                &" LEFT JOIN akt_typer aty ON (aty.aty_id = tfaktim) "_
                &" LEFT JOIN medarbejdere m ON (m.mid = tmnr) "_
                &" WHERE ("& aty_sql_realhours &") AND"_
				&" tjobnr = '"& LastJobID &"' AND ("& aktiviteterFundetpaFak &")"_
                &" AND tdato BETWEEN '"& lastIstDato &"' AND '"& lastIstDato2 &"'"_
                &" ORDER BY tdato, taktivitetid, tmnr" 
				
				'Response.write strSQLjoblogIF & " ("& lastFaknr &")<br><br>"
				'Response.flush
				
				oRec3.open strSQLjoblogIF, oConn, 3 
				jl = 0
				while not oRec3.EOF 
                    
                   
					
					'if jl = 0 then
					'strTxtExport = strTxtExport & vbcrlf
					'end if
					
					strTxtExport = strTxtExport & vbcrlf & "Joblog ikke faktureret;;"
					'strTxtExport = strTxtExport & lastFakStamoplys

                    jlidatoAAR = datepart("yyyy", oRec3("tdato"), 2,2) 
                    jlidatoMD = datepart("m", oRec3("tdato"), 2,2) 
                    strTxtExport = strTxtExport & replace(lastFakStamoplys, "#", oRec3("tdato")&";"& jlidatoMD &";"& jlidatoAAR)
					
					
                    
                    if isNull(oRec3("sumjoglogakttimer")) <> true then
                    jlogif_antalSigned = lastFaksign&((oRec3("sumjoglogakttimer"))/1)
                    else
                    jlogif_antalSigned = 0
                    end if
                    
                    

					'** Kurs på timereg SKAL benyttes til sammenligningen, **'
					'** Valuta omregnes altid til BasisValuta på udtræk ***'
                    if isNull(oRec3("timepris")) <> true then
                    jlogif_timeprisSigned = lastFaksign&((oRec3("timepris"))/1)
                    else
                    jlogif_timeprisSigned = 0
                    end if

                    if isNull(oRec3("kurs")) <> true then
                    ks = oRec3("kurs")
                    else
                    ks = 100
                    end if

                    call beregnValuta(jlogif_timeprisSigned,ks,100)
                    tpThis = formatnumber(valBelobBeregnet)

                    if isNull(oRec3("sumjoglogakttimer")) <> true AND isNull(oRec3("timepris")) <> true then
                    jlogBelob = lastFaksign&((oRec3("sumjoglogakttimer") * oRec3("timepris"))/1)  
					else
                    jlogBelob = 0
                    end if

					call beregnValuta(jlogBelob,ks,100)
                    jlogBelob = formatnumber(valBelobBeregnet)
					
                    

					strTxtExport = strTxtExport & oRec3("aktnavn") &";"
					strTxtExport = strTxtExport & oRec3("mnavn") &"; "&oRec3("mnr")&";"& oRec3("init")& ";"
					strTxtExport = strTxtExport & jlogif_antalSigned &";"
					strTxtExport = strTxtExport & tpThis &";"&basisValISO&";"
					
                    select case oRec3("aty_enh")
                    case 1
                    jlenh = "stk."
                    case 3
                    jlenh = "km"
                    case else
                    jlenh = "timer"
                    end select

					strTxtExport = strTxtExport &""&jlenh&";"

					
					strTxtExport = strTxtExport &""& jlogBelob &";"&basisValISO&";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport &";;"
					'strTxtExport = strTxtExport & vbcrlf
				
				jl = jl + 1	
				oRec3.movenext
				wend
				oRec3.close 

                end if

                strTxtExport = strTxtExport & vbcrlf 

    end sub


    call akttyper2009(2)

	call TimeOutVersion()
	
	    select case visning
		case 2
		ext = "txt"
		case else
		ext = "csv"
		end select
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\erp_fakturaer_eksport_2007.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& ext &"", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& ext &"", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& ext &"", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& ext &"", 8, -1)
		
	end if
	
	file = "fakturaexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& ext 
	
	
	
    'Response.write request("fakids")
    'Response.end
	
	fakids = request("fakids")
	fakIder = split(fakids, ",") 
	for i = 0 to Ubound(fakIder)
	
		if i = 0 then
		fakidKri = "(f.fid =  " & fakIder(i)
		else
		fakidKri = fakidKri & " OR f.fid =  " & fakIder(i)
		end if
	
	next
	
	fakidKri = fakidKri & ") AND shadowcopy <> 1"
	
	

    '************************************************************************
    '************************ HEADER + BODY *********************************
    '************************************************************************
	select case visning
	case 0 'Detaljer
	
	strSQL = "SELECT f.fid, f.faknr, f.dato AS fdato, "_
	&" f.fakdato, f.jobid, f.timer AS antal, f.beloeb, f.moms, "_
	&" f.kommentar, f.faktype, f.parentfak, f.aftaleid, f.enhedsang, f.rabat, f.jobbesk, f.fak_ski, f.istdato, f.istdato2, "_
    &" j.jobnavn, j.jobnr, j.jobknr, j.id AS jobid, "_
	&" k.Kkundenavn, k.Kkundenr, sk2.Kkundenavn AS sknavn, sk2.Kkundenr AS sknr, "_
	&" tidspunkt AS faktidspkt, f.betalt, f.faktype, f.betalt, f.valuta, f.kurs, "_
    &" ja1.mnavn AS jobans1, ja1.mnr AS jans1nr, ja2.mnavn AS jobans2, ja2.mnr AS jans2nr, "_
	& "fd.beskrivelse, fd.antal AS fakdet_antal, fd.enhedspris AS timepris, "_
	&" fd.aktpris AS fakdet_belob, fd.valuta AS fdvaluta, fd.kurs AS fdkurs, "_
	&" ka.mnavn AS kans, ka.mnr AS kansnr, "_ 
	&" ka2.mnavn AS kans2, ka2.mnr AS kansnr2,"_
	&" ska.mnavn AS skans, ska.mnr AS skansnr, "_
	&" ska2.mnavn AS skans2, ska2.mnr AS skansnr2,"_
	&" fd.aktid AS fdaktid, a.fase,"_
	& "a.navn AS aktnavn, fo.navn AS fomr, fd.enhedsang As fakdet_enhedsang, "_
	&" fd.rabat AS fakdet_rabat, v2.valutakode AS v2valutakode, medregnikkeioms, brugfakdatolabel, labeldato, fd.momsfri, f.fak_fomr, f.momssats, f.konto, f.modkonto, f.b_dato"_
	&" FROM fakturaer f "_
	&" LEFT JOIN job j ON (j.id = f.jobid)"_ 
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr)"_ 
	&" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid)"_ 
	&" LEFT JOIN medarbejdere ka ON (ka.mid = k.kundeans1)"_ 
	&" LEFT JOIN medarbejdere ka2 ON (ka.mid = k.kundeans2)"_ 
	&" LEFT JOIN medarbejdere ja1 ON (ja1.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere ja2 ON (ja2.mid = j.jobans2)"_
	&" LEFT JOIN aktiviteter a ON (a.id = fd.aktid)"_ 
	&" LEFT JOIN fomr fo ON (fo.id = a.fomr)"_ 
	&" LEFT JOIN serviceaft s ON (s.id = f.aftaleid)"_
	&" LEFT JOIN kunder sk2 ON (sk2.kid = s.kundeid)"_ 
	&" LEFT JOIN medarbejdere ska ON (ska.mid = sk2.kundeans1)"_ 
	&" LEFT JOIN medarbejdere ska2 ON (ska2.mid = sk2.kundeans2)"_
	&" LEFT JOIN valutaer v2 ON (v2.id = fd.valuta)"_
	&" WHERE "& fakidKri &" ORDER BY f.fakdato DESC, tidspunkt DESC"
	
	'Response.write strSQL & "<hr>"
	'Response.flush
	'Response.end
		
    '**** HEADER LINJE - Detaljeret ***
	strTxtExport = "Linietype;Aftale/Job;Kontakt;Kontakt id;Kontaktansv.1;Kontaktansv.1 nr.;"_
	&"Kontaktansv.2;Kontaktansv.2 nr.;Jobnavn;Jobnr;Jobansvarlig;Jobansv. nr.;"_
	&"jobejer;Jobejer nr.;Faktura nr.;Faktura beløb excl. moms; Moms; Faktura beløb incl. moms;Fakturadato (system) / Reg.dato;Måned;År;Periode Startdato;Periode Slutdato;Type;Kladde/Godkendt;"_
	&"Aktivitetsnavn;Medarb./Materiale Navn;Medarb./Vare Nr;Initialer;Antal;Enhedspris;Valuta;Enhed;Rabat;"_
	&"Beløb excl. moms;Valuta;Ventetimer brugt;Ventetimer ultimo;Vist på Faktura?;"_
	&"Forretningsområde;Intern;Medarb.linier findes?(sumakt/andre?);Fakturalinie tekst (20 kar.);"


         if instr(lto, "epi") <> 0 OR lto = "intranet - local" then 'Alle EPI
	     strTxtExport = strTxtExport &"Afdeling;Underafdeling;Momskode;Konto;Modkonto;Forfaldsdato;"
	     end if

	strTxtExport = strTxtExport & vbcrlf
	
	case 1,2 'Simpel
	
	strSQL = "SELECT f.fid, f.faknr, f.fakdato, f.timer, f.beloeb, f.b_dato, f.fakadr, f.faktype, "_
	&" f.konto, f.modkonto, f.moms, f.erfakbetalt, f.valuta, f.kurs, fak_ski, "_
	&" k.kkundenr, k.kkundenavn, k.adresse, k.postnr, k.city, k.land, k.telefon, f.valuta, "_
	&" f.kurs, k.ean, v.valutakode, k.cvr, medregnikkeioms, j.jobnr, s.aftalenr, brugfakdatolabel, labeldato, f.fak_fomr, k.ktype, momssats"_
	&" FROM fakturaer f "_
	&" LEFT JOIN kunder k ON (k.kid = f.fakadr)"_
    &" LEFT JOIN job j ON (j.id = f.jobid)"_
    &" LEFT JOIN serviceaft s ON (s.id = f.aftaleid)"_ 
	&" LEFT JOIN valutaer v ON (v.id = f.valuta)"_
	&" WHERE "& fakidKri &" ORDER BY f.fakdato DESC, tidspunkt DESC"
	
	'Response.Write strSQL
	'Response.flush
	
	    select case visning
	    case 1
	    'Adresse;Postnr;By;Land;
	    strTxtExport = "Kontakt;Kontakt id;"
	    
	    if lto = "acc" then
	    strTxtExport = strTxtExport & "Adresse;Postnr;By;Land;"
	    end if
	    
	    if lto <> "acc" then
	    strTxtExport = strTxtExport & "CVR;"
	    end if
	    
	    strTxtExport = strTxtExport & "Tlf;Faktura nr.;Fakturadato;F/L (Fak. systemdato:0, Labeldato:1);Beløb ekskl. moms "& basisValISO &";Moms;"
        
        strTxtExport = strTxtExport & "Type (0:Faktura, 1:Kreditnota);Konto;Modkonto;Forfaldsdato;Modtaget/Betalt;EAN;"
	    
        if lto <> "acc" then ' <> "acc" AND lto <> "jttek" AND lto <> "epi" then
            strTxtExport = strTxtExport & "Beløb inkl. moms "& basisValISO &";"
        end if
                    
	    if lto <> "acc" AND lto <> "synergi1" then '= "execon" OR lto = "immenso" OR lto = "intranet - local" then
	    strTxtExport = strTxtExport &"SKI;Kørsel ekskl. moms;Mat. forbrug / Udlæg ekskl. moms;Aktiviteter ekskl. moms;Intern;Valuta;Beløb ekskl. moms i faktura valuta;Jobnr;Aftale nr;"
	    end if

        if lto = "acc" then 'Accounting
	    strTxtExport = strTxtExport &"Udlæg (u.leverandør) ekskl. moms;Udlægs konto;Udlægs modkonto;"
	    end if


        if lto <> "acc" AND lto <> "synergi1" then
	    strTxtExport = strTxtExport &"Moms i faktura valuta;"
	    end if

        if instr(lto, "epi") <> 0 then 'Alle EPI
	    strTxtExport = strTxtExport &"Afdeling;Underafdeling;Momskode;"
	    end if
	                
	    strTxtExport = strTxtExport & vbcrlf
	    case 2
	        if lto <> "execon" AND lto <> "immenso" then
	        strTxtExport = "Debitornr.;Fakturanr.;Fakturadato;Forfaldsdato;Beløb (incl. moms);Valuta;"
	        strTxtExport = strTxtExport & vbcrlf
	        end if
	    end select
	
	end select
	lastFaknr = 0
    f = 0
    
    
    'response.write strSQL
    'response.flush
    
    oRec.open strSQL, oConn, 3
	while not oRec.EOF 
    
    FakKurs = oRec("kurs")
    
    if len(oRec("fid")) <> 0 then
	fakturaidThis = oRec("fid")
	else
	fakturaidThis = 0
	end if
				
    
    select case visning
	case 0 '*** DETALJER
     
    
     
    
                    '*** Faktura på aftale ell. job ***
                    if oRec("aftaleid") <> 0 then
                    knavn = oRec("sknavn") 
					knr = oRec("sknr") 
					kans1 = oRec("skans") 
					kans1nr = oRec("skansnr")
					kans2 = oRec("skans2") 
					kans2nr = oRec("skansnr2")
					
					else
                    knavn = oRec("kkundenavn") 
					knr = oRec("kkundenr") 
					kans1 = oRec("kans") 
					kans1nr = oRec("kansnr")
					kans2 = oRec("kans2") 
					kans2nr = oRec("kansnr2")
				

				    end if
					
					jobnavn = oRec("jobnavn") 
					jobnr = oRec("jobnr") 
					jobans1 = oRec("jobans1")
					jobans1nr = oRec("jans1nr")
					jobans2 = oRec("jobans2")
					jobans2nr = oRec("jans2nr")
					faknr = oRec("faknr") 

                    if instr(lto, "epi") <> 0 then
                     
                        if oRec("brugfakdatolabel") = 1 then
                        fakdato = formatdatetime(oRec("labeldato")) 
                        else
                        fakdato = formatdatetime(oRec("fakdato"))
                        end if
        
                    else
                    '*** Altid faktura system dato pga. forbindelse til omsætnings måned
					fakdato = formatdatetime(oRec("fakdato"))
                    end if            


                    fakdatoAAR = datepart("yyyy", fakdato, 2,2) 
                    fakdatoMD = datepart("m", fakdato, 2,2) 
                    'end if
					betalt = oRec("betalt")
                    
                    fakistdato = formatdatetime(oRec("istdato"))
                    fakistdato2 = formatdatetime(oRec("istdato2"))

                    fakBelobexMoms = formatnumber(oRec("beloeb"),2)
                    fakMoms = formatnumber(oRec("moms"),2)
                    fakBelobinclMoms = formatnumber((oRec("beloeb")+oRec("moms")),2)
                   

                    '*** Main Faktura linie ****'
                    if lastFakID <> oRec("fid") then
                    aktiviteterFundetpaFak = " taktivitetid <> 0 "
                        
                        '** Materialer Udlæg og job på aktivitet der ikke er med på faktura **'
                        if f <> 0 then
                        call joblogIFog_udlag
                        end if
                    
                    strTxtExportStam = ""    
                    '*** Export ****
                            
                             

                            strTxtExport = strTxtExport & vbcrlf
		                    strTxtExport = strTxtExport & "Faktura;"

                             if oRec("aftaleid") <> 0 then
                             strTxtExport = strTxtExport & "Aftale;"
					         else
                             strTxtExport = strTxtExport & "Job;"
					         end if

		                    strTxtExportStam = strTxtExportStam & knavn &";"
		                    strTxtExportStam = strTxtExportStam & knr &";"
		                    strTxtExportStam = strTxtExportStam & kans1 &";"
		                    strTxtExportStam = strTxtExportStam & kans1nr &";"
		                    strTxtExportStam = strTxtExportStam & kans2 & ";"
		                    strTxtExportStam = strTxtExportStam & kansnr2 &";"
		                    strTxtExportStam = strTxtExportStam & oRec("jobnavn") &";"
		                    strTxtExportStam = strTxtExportStam & oRec("jobnr") &";"
		                    strTxtExportStam = strTxtExportStam & oRec("jobans1")&";"
		                    strTxtExportStam = strTxtExportStam & oRec("jans1nr")&";"
		                    strTxtExportStam = strTxtExportStam & oRec("jobans2")&";"
		                    strTxtExportStam = strTxtExportStam & oRec("jans2nr")&";"
		                    strTxtExportStam = strTxtExportStam & oRec("faknr") &";"

                            faksign = ""
                            faktypeThis = "-"
		                    select case oRec("faktype")
		                    case 1
		                    faktypeThis = "Kreditnota"
                            faksign = "-"
		                    case 2
		                    faktypeThis = "Rykker"
                            faksign = ""
		                    case else
		                    faktypeThis = "Faktura"
                            faksign = ""
		                    end select  

                            strTxtExport = strTxtExport & strTxtExportStam

                            strTxtExport = strTxtExport & faksign&(fakBelobexMoms) &";"
                            strTxtExport = strTxtExport & faksign&(fakMoms) &";"
                            strTxtExport = strTxtExport & faksign&(fakBelobinclMoms) &";"
        
		                    strTxtExport = strTxtExport & fakdato &";" & fakdatoMD &";"& fakdatoAAR &";"
                            strTxtExportStam = strTxtExportStam &";;;#;" '"& fakdato &";"

                            strTxtExport = strTxtExport & fakistdato &";"
                            strTxtExport = strTxtExport & fakistdato2 &";"
		
		                  
		
		
		                    strTxtExport = strTxtExport & faktypeThis &";"
                            strTxtExportStam = strTxtExportStam &";;"& faktypeThis &";;"

		                    strTxtExport = strTxtExport & oRec("betalt") &";"

                            lastFakmoplyStas = strTxtExportStam
		
		                    'strTxtExportStam = strTxtExportStam & ";"
		
		                    'strTxtExportStam = strTxtExportStam & ";;"
		                    'strTxtExportStam = strTxtExportStam & ";"

                            if instr(lto, "epi") <> 0 OR lto = "intranet - local" then 'Alle EPI forfaldsto med på alle linjer
                            strTxtExport = strTxtExport &";;;;;;;;;;;;;;;;;;;;;;;"& oRec("b_dato") &";"
                            end if

                            

                            

                    end if


		
		'*** Export ****
         '*** Faktura på aftale ell. job ***
                   
		strTxtExport = strTxtExport & vbcrlf
		strTxtExport = strTxtExport & "Fakturalinie;;"
       			
	    strTxtExport = strTxtExport & knavn &";"
		strTxtExport = strTxtExport & knr &";"
		strTxtExport = strTxtExport & kans1 &";"
		strTxtExport = strTxtExport & kans1nr &";"
		strTxtExport = strTxtExport & kans2 & ";"
		strTxtExport = strTxtExport & kansnr2 &";"
		strTxtExport = strTxtExport & oRec("jobnavn") &";"
		strTxtExport = strTxtExport & oRec("jobnr") &";"
		strTxtExport = strTxtExport & oRec("jobans1")&";"
		strTxtExport = strTxtExport & oRec("jans1nr")&";"
		strTxtExport = strTxtExport & oRec("jobans2")&";"
		strTxtExport = strTxtExport & oRec("jans2nr")&";"
		strTxtExport = strTxtExport & oRec("faknr") &";"

        strTxtExport = strTxtExport &";"
        strTxtExport = strTxtExport &";"
        strTxtExport = strTxtExport &";"
        
		strTxtExport = strTxtExport & fakdato &";" & fakdatoMD &";"& fakdatoAAR &";"

        strTxtExport = strTxtExport &";"
        strTxtExport = strTxtExport &";"
		
		
		strTxtExport = strTxtExport & faktypeThis &";"
		strTxtExport = strTxtExport & ";"
		
		strTxtExport = strTxtExport & oRec("aktnavn") &";"
		
		strTxtExport = strTxtExport & ";;"
		strTxtExport = strTxtExport & ";"
		
       
        'Response.Write "faksign:" & faksign
        'Response.flush
        timeprisSigned = faksign&((oRec("timepris"))/1)


		if oRec("aftaleid") <> 0 then
		    
            antalSigned = faksign&((oRec("antal"))/1)
            belobSigned = faksign&((oRec("beloeb"))/1)

		    strTxtExport = strTxtExport & antalSigned &";"
		    
		    enhpris = 0
		    if oRec("antal") <> 0 then
		    enhpris = formatnumber((belobSigned/antalSigned), 2)
		    else
		    stkpris = belobSigned
		    end if
		    
		    strTxtExport = strTxtExport & enhpris &";"& oRec("v2valutakode") &";"
		    strTxtExport = strTxtExport &"Enheder;"
		    strTxtExport = strTxtExport & oRec("rabat") &";"
		    
            if len(trim(belobSigned)) <> 0 then
            belobSigned = belobSigned
            else
            belobSigned = 0
            end if

		    call beregnValuta(belobSigned,FakKurs,100)
		    strTxtExport = strTxtExport &formatnumber(valBelobBeregnet) &";"& basisValISO &";"
		    
		
		else
		    
            if len(trim(oRec("timepris"))) <> 0 then
            tpThisAkt = formatnumber(timeprisSigned)
            else
            tpThisAkt = 0
            end if

            antalSigned = faksign&((oRec("fakdet_antal"))/1)
            belobSigned = faksign&((oRec("fakdet_belob"))/1)

		    strTxtExport = strTxtExport & antalSigned &";"
		    strTxtExport = strTxtExport & tpThisAkt &";"
		    strTxtExport = strTxtExport & oRec("v2valutakode") &";"
    		
    		
		    strEnhedfd = ""
            select case oRec("fakdet_enhedsang")
		    case 0
		    strEnhedfd = "timer"
		    case 1
		    strEnhedfd = "stk."
		    case 2
		    strEnhedfd = "enheder"
            case 3
		    strEnhedfd = "km"
            case else
		    strEnhedfd = ""
            end select
    		
    		
		    strTxtExport = strTxtExport & strEnhedfd &";"
		    strTxtExport = strTxtExport & oRec("fakdet_rabat") &";"
		    
		    if len(trim(belobSigned)) <> 0 then
            belobSigned = belobSigned
            else
            belobSigned = 0
            end if
		    
		    call beregnValuta(belobSigned, FakKurs, 100)
		    strTxtExport = strTxtExport & formatnumber(valBelobBeregnet) &";"& basisValISO &";"
    		
		end if
		
		
		strTxtExport = strTxtExport & ";"
		strTxtExport = strTxtExport & ";"
		strTxtExport = strTxtExport & "1;"
		strTxtExport = strTxtExport & oRec("fomr") &";"
        strTxtExport = strTxtExport & oRec("medregnikkeioms") &";"
        
		
		
		
		
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
				
				
				'*** Valuta ****'
				if len(trim(oRec("fdvaluta"))) <> 0 then
				valutaAktID = oRec("fdvaluta")
				else
				valutaAktID = 1
				end if
				



                '*********** MEDARBLINER OG JOBLOG **********************

                if instr(lto, "epi") <> 0 OR lto = "intranet - local" then
                medablinjerOn = 0
                else
                medablinjerOn = 1
                end if


                if cint(medablinjerOn) = 1 then


				'** Medarbejdere ***'
				strSQL2 = "SELECT fms.fak AS fmstimer, fms.showonfak AS fmsshowonfak,"_
				&" fms.enhedspris AS fmsenhedspris, fms.beloeb AS fmsbelob, fms.tekst, m.mid, "_
				&" m.mnavn, m.mnr, m.init, fms.venter AS fmsventer, fms.venter_brugt AS fmsventebrugt, "_
				&" enhedsang, medrabat, fms.valuta, fms.kurs, v3.valutakode AS v3valutakode "_
				&" FROM fak_med_spec fms "_
				&" LEFT JOIN medarbejdere m ON (m.mid = fms.mid) "_
				&" LEFT JOIN valutaer v3 ON (v3.id = fms.valuta)"_
				&" WHERE fms.fakid = "& fakturaidThis &" AND fms.aktid = "& aktidThis &" AND fms.enhedspris = " & tprisThis & " AND fms.valuta = "& valutaAktID 
				
				'Response.write strSQL2
				'Response.flush
				
				oRec2.open strSQL2, oConn, 3 
				m = 0
				while not oRec2.EOF 
					
					if cint(m) = 0 then
					strTxtExport = strTxtExport & "1;"

                    strTXTthis = oRec("beskrivelse")

                  
                    strTXTthis = left(strTXTthis, 20) & ""
                    strTxtExport = strTxtExport & Chr(34) & trim(strTXTthis) & Chr(34) & ";" 

					
					'strTxtExport = strTxtExport & vbcrlf
					end if
					
					strTxtExport = strTxtExport & vbcrlf & "Medarb.;;"
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

                    strTxtExport = strTxtExport &";"
                    strTxtExport = strTxtExport &";"
                    strTxtExport = strTxtExport &";"

					strTxtExport = strTxtExport & fakdato &";" & fakdatoMD &";"& fakdatoAAR &";"
                    strTxtExport = strTxtExport &";"
                    strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport & faktypeThis &";"

                    strTxtExport = strTxtExport &";"
					strTxtExport = strTxtExport & oRec("aktnavn") &";"
					
					strTxtExport = strTxtExport & oRec2("mnavn") &";"
					strTxtExport = strTxtExport & oRec2("mnr")&";"
					strTxtExport = strTxtExport & oRec2("init")&";"

                    fms_enhprisSigned = faksign&((oRec2("fmsenhedspris"))/1)
                    fms_antalSigned = faksign&((oRec2("fmstimer"))/1)
                    fms_belobSigned = faksign&((oRec2("fmsbelob"))/1)

					strTxtExport = strTxtExport & fms_antalSigned &";"
					strTxtExport = strTxtExport & formatnumber(fms_enhprisSigned) &";"
					
                    strTxtExport = strTxtExport & oRec2("v3valutakode") &";"
					
					strEnhed = ""
					select case oRec2("enhedsang")
					case 0
					strEnhed = "timer"
					case 1
					strEnhed = "stk."
					case 2
					strEnhed = "enheder"
                    case 3
					strEnhed = "km"
                    case else
                    strEnhed = ""
					end select
					
					strTxtExport = strTxtExport & strEnhed &";"
					strTxtExport = strTxtExport & oRec2("medrabat") &";"
					
					'** Kurs på Faktura SKAL benyttes til sammenligningen, **'
					'** da sumbeløbet på medabl linie er beregnet udfra **'
					'** kurs på faktura ***'
					if len(trim(fms_belobSigned)) <> 0 then
                    fms_belobSigned = fms_belobSigned
                    else
                    fms_belobSigned = 0
                    end if

					call beregnValuta(fms_belobSigned,FakKurs,100)
					
					strTxtExport = strTxtExport & formatnumber(valBelobBeregnet) &";"& basisValISO &";"
					strTxtExport = strTxtExport & oRec2("fmsventebrugt") &";"
					strTxtExport = strTxtExport & oRec2("fmsventer") &";"
					strTxtExport = strTxtExport & oRec2("fmsshowonfak") &";"
					strTxtExport = strTxtExport & oRec("fomr") &";"
					strTxtExport = strTxtExport & ";;"

                    strTXTthis = oRec2("tekst") 

                   

                     strTXTthis = left(strTXTthis, 20) & ""
                     strTxtExport = strTxtExport & Chr(34) & trim(strTXTthis) & Chr(34) & ";" 


					
                
                        if instr(lto, "epi") <> 0 OR lto = "intranet - local" then
                        joblogOn = 0
                        else
                        joblogOn = 1
                        end if


                        if cint(joblogOn) = 1 then

                    

                            
                                        '** Joblog i fakturaperiode på aktivitet ***

                                        if isNull(oRec2("mid")) <> true then
                                        thisMid = oRec2("mid")
                                        else
                                        thisMid = 0
                                        end if

				                        strSQLjoblog = "SELECT timer AS sumjoglogakttimer, tmnr, tjobnr, taktivitetid, timepris,"_
                                        &" a.navn AS aktnavn, m.mnavn, m.mnr, m.init, t.valuta, t.kurs, a.id AS aktid, tdato, aty.aty_enh "_
                                        &" FROM timer AS t"_
                                        &" LEFT JOIN aktiviteter AS a ON (a.id = taktivitetid) "_
                                        &" LEFT JOIN akt_typer AS aty ON (aty.aty_id = tfaktim) "_
                                        &" LEFT JOIN medarbejdere AS m ON (m.mid = tmnr) "_
                                        &" WHERE "_
				                        &" taktivitetid = "& aktidThis &" AND tmnr = "& thisMid &""_
                                        &" AND tdato BETWEEN '"& year(oRec("istdato")) &"/"& month(oRec("istdato")) &"/"& day(oRec("istdato")) &"' AND '"& year(oRec("istdato2")) &"/"& month(oRec("istdato2")) &"/"& day(oRec("istdato2")) &"'"_
                                        &" ORDER BY tdato, taktivitetid, tmnr" 
				
				                        'Response.write strSQLjoblog
				                        'Response.end
				
               

				                        oRec3.open strSQLjoblog, oConn, 3 
				                        jl = 0
				                        while not oRec3.EOF 
                    
                                            aktiviteterFundetpaFak = aktiviteterFundetpaFak & " AND taktivitetid <> "& oRec3("aktid")  
					
					                        'if jl = 0 then
					                        'strTxtExport = strTxtExport & vbcrlf
					                        'end if
					
					                        strTxtExport = strTxtExport & vbcrlf & "Joblog;;"
					                        strTxtExport = strTxtExport & knavn &";"
					                        strTxtExport = strTxtExport & knr &";"
					                        strTxtExport = strTxtExport & kans1 &";"
					                        strTxtExport = strTxtExport & kans1nr  &";"
					                        strTxtExport = strTxtExport & kans2 &";"
					                        strTxtExport = strTxtExport & kansnr2 &";"
					                        strTxtExport = strTxtExport & jobnavn &";"
					                        strTxtExport = strTxtExport & jobnr &";"
					                        strTxtExport = strTxtExport & jobans1 &";"
					                        strTxtExport = strTxtExport & jobans1nr &";"
					                        strTxtExport = strTxtExport & jobans2 &";"
					                        strTxtExport = strTxtExport & jobans2nr &";"
					
                                            strTxtExport = strTxtExport & faknr &";"
                                            strTxtExport = strTxtExport &";"
                                            strTxtExport = strTxtExport &";"
                                            strTxtExport = strTxtExport &";"

                                            jldatoAAR = datepart("yyyy", oRec3("tdato"), 2,2) 
                                            jldatoMD = datepart("m", oRec3("tdato"), 2,2) 

					                        strTxtExport = strTxtExport & oRec3("tdato") &";"& jldatoMD &";"& jldatoAAR &";" 'fakdato
                                            strTxtExport = strTxtExport &";"
                                            strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport & faktypeThis &";"
					                        strTxtExport = strTxtExport &";"
					
					
					
					                        '** Kurs på timereg SKAL benyttes til sammenligningen, **'
					                        '** Valuta omregnes altid til basisValISO på udtræk ***'
                                            jlog_timeprisSigned = faksign&((oRec3("timepris"))/1)
                    
                                            if len(trim(jlog_timeprisSigned)) <> 0 then
                                            jlog_timeprisSigned = jlog_timeprisSigned
                                            else
                                            jlog_timeprisSigned = 0
                                            end if

                                            call beregnValuta(jlog_timeprisSigned,oRec3("kurs"),100)
                                            tpThis = formatnumber(valBelobBeregnet)

                                            jlogBelob = faksign&((oRec3("sumjoglogakttimer") * oRec3("timepris"))/1) 
                    
                                            if len(trim(jlogBelob)) <> 0 then
                                            jlogBelob = jlogBelob
                                            else
                                            jlogBelob = 0
                                            end if
                     
					                        jlog_antalSigned = faksign&((oRec3("sumjoglogakttimer"))/1)
                   
                                            call beregnValuta(jlogBelob,oRec3("kurs"),100)
                                            jlogBelob = formatnumber(valBelobBeregnet)
					
					                        strTxtExport = strTxtExport & oRec3("aktnavn") &";"
					                        strTxtExport = strTxtExport & oRec3("mnavn") &"; "&oRec3("mnr")&";"& oRec3("init")& ";"
					                        strTxtExport = strTxtExport & jlog_antalSigned &";"
					                        strTxtExport = strTxtExport & tpThis &";"& basisValISO &";"
					
                                            select case oRec3("aty_enh")
                                            case 1
                                            jlenh = "stk."
                                            case 3
                                            jlenh = "km"
                                            case else
                                            jlenh = "timer"
                                            end select

					                        strTxtExport = strTxtExport &""&jlenh&";"
					                        strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport &""& jlogBelob &";"& basisValISO &";"
					                        strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport &";"
					                        strTxtExport = strTxtExport &";;"
					                        'strTxtExport = strTxtExport & vbcrlf
				
				                        jl = jl + 1	
				                        oRec3.movenext
				                        wend
				                        oRec3.close 

                                        strTxtExport = strTxtExport & vbcrlf
				                        '**** END JOBLOG *******************************

                                        end if 'joblogOn
				
				m = m + 1	
				oRec2.movenext
				wend
				oRec2.close 

            
                end if 'medablinjerJoblogOn




				
				if m = 0 then
				strTxtExport = strTxtExport & "0;"
				
				if oRec("aftaleid") <> 0 then
                strTXTthis = oRec("jobbesk")
				else
                strTXTthis = oRec("beskrivelse")
				end if
				
              
                     strTXTthis = left(strTXTthis, 20) & ""
                     strTxtExport = strTxtExport & Chr(34) & trim(strTXTthis) & Chr(34) & ";" 


				end if



                    if instr(lto, "epi") <> 0 OR lto = "intranet - local" then 'Alle EPI
            
                        fak_fomrTxt = "" 'UnderAFD
                        fak_AfdTxt = ""
                        'ktypeTxt = "" 'AFD
                        if len(trim(oRec("fak_fomr"))) <> 0 then
                        fak_fomr = oRec("fak_fomr")
                        else
                        fak_fomr = 0
                        end if
            
                          

                                         '*** FOMR = UnderAFD, UNIT = AFD ****'
                                         strSQLfomr = "SELECT navn, business_area_label, business_unit FROM fomr WHERE id = " & fak_fomr
                
                                         oRec4.open strSQLfomr, oConn, 3
							             if not oRec4.EOF then
							 
							             fak_fomrTxt = oRec4("business_area_label")
                                         fak_AfdTxt = oRec4("business_unit")
							    
							             end if
							             oRec4.close


                            if isNull(oRec("momsfri")) <> true then

                            if cint(oRec("momsfri")) <> 1 then
                            momssats = oRec("momssats")
                            else
                            momssats = 0
                            end if
        
                            else

                            momssats = 0

                            end if

                        konto = oRec("konto")
                        modKonto = oRec("modkonto")
                        


	                    strTxtExport = strTxtExport &""& fak_AfdTxt &";"& fak_fomrTxt &";"& momssats &";"& konto &";"& modKonto &";"& oRec("b_dato") &";"

	                end if


        
               
        '*** LOOP VALUES                        
        if oRec("aftaleid") <> 0 then
            lastFakType = "aftale"
	    else
            lastFakType = "job"
        end if

              
              
    
        lastFaknr = oRec("faknr")
        lastJobID = oRec("jobid")
        lastIstDato = year(oRec("istdato")) &"/"& month(oRec("istdato")) &"/"& day(oRec("istdato"))
        lastIstDato2 = year(oRec("istdato2")) &"/"& month(oRec("istdato2")) &"/"& day(oRec("istdato2"))
        lastFakID = oRec("fid")
        lastFakKurs = FakKurs
        lastFakStamoplys = strTxtExportStam
        lastFakDato = fakdato
        lastFakSign = faksign

        '******************************



				
    case 1 'Standard
    
    if len(trim(oRec("beloeb"))) <> 0 then
         if oRec("faktype") = 0 then
        Belob_fakvaluta = oRec("beloeb")
        moms_fakvaluta = oRec("moms")
        else
        Belob_fakvaluta = -(oRec("beloeb"))
        moms_fakvaluta = -(oRec("moms"))
        end if
    else
    Belob_fakvaluta = 0
    moms_fakvaluta = 0
    end if

    if oRec("valuta") <> 1 then 
            
            'if oRec("beloeb") <> "" then
            'belobsigned = 
            call beregnValuta(oRec("beloeb"),oRec("kurs"),100)
            fakBelob = valBelobBeregnet

            call beregnValuta(oRec("moms"),oRec("kurs"),100)
            fakMoms = valBelobBeregnet

            
    else
        
        fakBelob = oRec("beloeb")
        fakMoms = oRec("moms")
    
    end if
    
    if oRec("faktype") = 0 then
    belob = fakBelob
    antal = oRec("timer")
    moms = fakMoms
    else 'Kreditnota

       
            
        belob = -(fakBelob) 
        moms = -(fakMoms) 

        antal = -(oRec("timer"))

        if cdbl(belob) > 0 then 'skirer fortegen er negtativ på kreditnota også ved fejl indtastniing af n.feks negativ timepris
        belob = belob * -1
        moms = moms * -1
        end if
   
    end if
    
    if len(trim(oRec("adresse"))) <> 0 then
    call htmlparseCSV(oRec("adresse"))
    adrParsed = htmlparseCSVtxt
    else
    adrParsed = ""
    end if
    
    strTxtExport = strTxtExport & vbcrlf
    strTxtExport = strTxtExport & oRec("kkundenavn") &";"& oRec("kkundenr") &";"
    
    if lto = "acc" then
    strTxtExport = strTxtExport & adrParsed &";"& oRec("postnr") &";"& oRec("city") &"; "& oRec("land")&";" 
    end if
    
    if lto <> "acc" then
    strTxtExport = strTxtExport & oRec("cvr") &";"
    end if
    
    strTxtExport = strTxtExport & oRec("telefon")&";"
    strTxtExport = strTxtExport & oRec("faknr")&";"
    
    if cint(oRec("brugfakdatolabel")) = 1 then
    strTxtExport = strTxtExport & oRec("labeldato")&";"& oRec("brugfakdatolabel") &";"
    else
    strTxtExport = strTxtExport & oRec("fakdato")&";"& oRec("brugfakdatolabel") &";"
    end if

    strTxtExport = strTxtExport & belob &";"& moms &";"& oRec("faktype")&";"

    if lto <> "synergi1" then
    konto = oRec("konto")
    else
        if len(trim(oRec("kkundenr"))) <> 0 AND isNull(oRec("kkundenr")) <> true then
        konto = replace(trim(oRec("kkundenr")), " ", "")
        else
        konto = 0
        end if
    end if

    modKonto = oRec("modkonto")

	strTxtExport = strTxtExport & konto &";"& modKonto &";"& oRec("b_dato") &";"&oRec("erfakbetalt")&";"& oRec("ean")&";"
    
    
    if lto <> "acc" then
    fakbelobInklMoms = (belob/1)+(moms/1)
    'fakbelobInklMoms = replace(fakbelobInklMoms, ",", "")
    'fakbelobInklMoms = replace(fakbelobInklMoms, ".", "")
    strTxtExport = strTxtExport & fakbelobInklMoms &";"
    end if       
            
            if lto <> "acc" AND lto <> "synergi1" then 'if lto = "execon" OR lto = "immenso" OR lto = "intranet - local" then
            
            strTxtExport = strTxtExport & oRec("fak_ski")&";"                 
                             
                             '** Kørsel Km ex. moms
		                     transbeloebPos = 0
		                         strSQLselmf = "SELECT SUM(aktpris*(kurs/100)) AS transbeloeb FROM faktura_det WHERE fakid = " & fakturaidThis & " AND enhedsang = 3 GROUP BY fakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 transbeloebPos = oRec4("transbeloeb")
							    
							 end if
							 oRec4.close
                             strTxtExport = strTxtExport & transbeloebPos & ";"
                             
                             '** Udlæg ex. moms
		                     matbeloebPos = 0
		                     strSQLselmf = "SELECT SUM(matbeloeb*(kurs/100)) AS matbeloeb FROM fak_mat_spec WHERE matfakid = " & fakturaidThis & " GROUP BY matfakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 matbeloebPos = oRec4("matbeloeb")
							    
							 end if
							 oRec4.close
                             strTxtExport = strTxtExport & matbeloebPos & ";"
                             
                             '** Aktiviteter ex. moms
		                     aktbeloebPos = 0
		                     strSQLselmf = "SELECT SUM(aktpris*(kurs/100)) AS aktbeloeb FROM faktura_det WHERE fakid = " & fakturaidThis & " AND enhedsang <> 3 GROUP BY fakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 aktbeloebPos = oRec4("aktbeloeb")
							    
							 end if
							 oRec4.close
                             
                             
            strTxtExport = strTxtExport & aktbeloebPos & ";" & oRec("medregnikkeioms") & ";"
            strTxtExport = strTxtExport & oRec("valutakode") &";"
            strTxtExport = strTxtExport & Belob_fakvaluta &";"
        
            strTxtExport = strTxtExport & oRec("jobnr") &";"
            strTxtExport = strTxtExport & oRec("aftalenr") &";"

          
          
        
            
            end if

            

            if lto = "acc" then
                
                 
	                        '** Udlæg ex. moms
		                     matbeloebPos = 0
		                     strSQLselmf = "SELECT SUM(matbeloeb*(kurs/100)) AS matbeloeb FROM fak_mat_spec WHERE matfakid = " & fakturaidThis & " GROUP BY matfakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 matbeloebPos = oRec4("matbeloeb")
							    
							 end if
							 oRec4.close
                             strTxtExport = strTxtExport & matbeloebPos & ";115;116;"
	            

            end if


            

            if lto <> "acc" AND lto <> "synergi1" then
            strTxtExport = strTxtExport & moms_fakvaluta &";"
            end if


            if instr(lto, "epi") <> 0 then 'Alle EPI
            
            fak_fomrTxt = "" 'UnderAFD
            fak_AfdTxt = ""
            'ktypeTxt = "" 'AFD
            if len(trim(oRec("fak_fomr"))) <> 0 then
            fak_fomr = oRec("fak_fomr")
            else
            fak_fomr = 0
            end if
            
                          

                             '*** FOMR = UnderAFD, UNIT = AFD ****'
                             strSQLfomr = "SELECT navn, business_area_label, business_unit FROM fomr WHERE id = " & fak_fomr
                
                             oRec4.open strSQLfomr, oConn, 3
							 if not oRec4.EOF then
							 
							 fak_fomrTxt = oRec4("business_area_label")
                             fak_AfdTxt = oRec4("business_unit")
							    
							 end if
							 oRec4.close

            momssats = oRec("momssats")


	        strTxtExport = strTxtExport &""& fak_AfdTxt &";"& fak_fomrTxt &";"& momssats &";"
	        end if

            
    case 2
    
    
    '** Omregner ikke til DKK valuta men agiver valuta på hver linie **'
     fakBelob = oRec("beloeb")
    
    if oRec("faktype") = 0 then
    belob = fakBelob
    antal = oRec("timer")
    moms = oRec("moms")
    else
    belob = -(fakBelob)
    antal = -(oRec("timer"))
    moms = -(oRec("moms"))
    end if
    
   
    if len(belob) <> 0 then
    belob = formatnumber((belob + moms), 2)
    belob = replace(belob, ",", "")
    belob = replace(belob, ".", "")
    else
    belob = 0
    end if
    
    
    fakDato = formatdatetime(oRec("fakdato"), 2)
    fakDato = replace(fakDato, "-", "")
    fakDato = left(fakDato,4) & right(fakDato, 2)
    
    fakforfaldDato = formatdatetime(oRec("b_dato"), 2)
    fakforfaldDato = replace(fakforfaldDato, "-", "")
    fakforfaldDato = left(fakforfaldDato,4) & right(fakforfaldDato,2)
    
    
    strTxtExport = strTxtExport & oRec("kkundenr") &","
    strTxtExport = strTxtExport & oRec("faknr")&","""& fakDato &""","""& fakforfaldDato &""","
    strTxtExport = strTxtExport & belob &","""& oRec("valutakode") &""""
	strTxtExport = strTxtExport & vbcrlf 
    
    
   
    
    end select		
	
    

                
    f = f + 1			
    oRec.movenext
	wend
	oRec.close
	

    select case visning
	case 0

    '*** Joblog Aktiviteter ikke med på faktura + Udlæg og materialer
    call joblogIFog_udlag

	
	
	
	          
				

                


	
	end select
	
	
	
	'strTxtExport = Replace(strTxtExport, "ø", "&oslash;")
	'strTxtExport = Replace(strTxtExport, "Ø", "&Oslash;")
	'strTxtExport = Replace(strTxtExport, "æ", "&aelig;")
	'strTxtExport = Replace(strTxtExport, "Æ", "&AElig;")
	'strTxtExport = Replace(strTxtExport, "å", "&aring;")
	'strTxtExport = Replace(strTxtExport, "Å", "&Aring;")
	
	
	objF.WriteLine(strTxtExport)
	objF.close


    '*** nulstiller LCID Bl.a EPI_UK

    sprog = 1 'DK
    if len(trim(session("mid"))) <> 0 then
    strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

    oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    sprog = oRec("sprog")
    end if
    oRec.close
    end if


    select case sprog
    case 1
    Session.LCID = 1030
    case 2
    Session.LCID = 2057
    case 3
    Session.LCID = 1053
    case else
    Session.LCID = 1030
    end select

	Response.redirect "../inc/log/data/"& file &""	
	
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
		
