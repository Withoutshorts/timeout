<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<!--include file="../inc/regular/header_hvd_inc.asp"-->

<%

if len(trim(session("user"))) = 0 then
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else

    fakid = request("fakid")
    
     
    function SQLBless2(s)
    dim tmp2
    tmp2 = s
    tmp2 = replace(tmp2, ",", ".")
    SQLBless2 = tmp2
    end function 
    
    function SQLBless3(s)
	dim tmp3
	tmp3 = s
	tmp3 = replace(tmp3, ".", ",")
	SQLBless3 = tmp3
	end function
    
    momskonto = 1 
    
    strSQL = "SELECT f.fid, f.jobid, f.editor, f.faknr, f.fakdato, f.beloeb, f.faktype, "_
    &" f.konto, f.modkonto, f.moms, j.jobnavn, f.jobbesk, f.aftaleid, f.subtotaltilmoms, f.momskonto, f.kurs, f.moms, f.medregnikkeioms, f.fak_laast FROM fakturaer f "_
    &" LEFT JOIN job j ON j.id = f.jobid WHERE fid = " & fakid
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then
        
        'varBilag = oRec("faknr")
        intBeloeb = oRec("beloeb")
        posteringsdato = year(oRec("fakdato")) &"/"& month(oRec("fakdato")) &"/"& day(oRec("fakdato"))
        strEditor = oRec("editor") 
        thisfakid = oRec("fid")
        intkontonr = oRec("konto")
        modkontonr = oRec("modkonto")
        intType = oRec("faktype")
        
        if oRec("aftaleid") <> 0 then
        posteringsTxt = left(oRec("jobbesk"), 20)
        else
        posteringsTxt = left(oRec("jobnavn"), 20)
        end if
        
        intMoms = oRec("moms")
        
        subtotaltilmoms = oRec("subtotaltilmoms")
        
        momskonto = oRec("momskonto")
        kurs = oRec("kurs")

        medregnikkeioms = oRec("medregnikkeioms")
        medregnikkeioms_opr = medregnikkeioms '*** Skal altid være identisk her pga faknr

        faknr = oRec("faknr")
        fak_laast = oRec("fak_laast")
    
    end if
    oRec.close
    
    debkre = "fak"
    strTekst = posteringsTxt
    vorref = session("Mid")
	intStatus = 1	
	'strDato = year(now) &"/"& month(now) &"/"& day(now)	
	
	showfakDato = posteringsdato

    '***** Status og faktura nr ****'
    '** finder faknr  ***'
    intFakbetalt = 1
    func = "dbred_gk"
    id = fakid

    '*** Hvis faktura har været godkedt skal den ikke kunne ændre faktura nr. **'
    if cint(fak_laast) <> 1 then
        call findFaknr(func)
    else
        intFaknum = faknr
    end if

    varBilag = intFaknum
	
    '*** Set faktura til godkendt **'
    '*** Opdaterer/Indsætter posterings ID på faktura ***' 
    strSQL = "UPDATE fakturaer SET betalt = 1, oprid = 0, faknr = "& intFaknum &", fak_laast = 1 WHERE fid = " & fakid
    oConn.execute(strSQL)


    '*** UDPATE også shadowcopy  ****
    select case lto
    case "nt"
    strSQL = "UPDATE fakturaer SET betalt = 1, oprid = 0, faknr = "& intFaknum &", fak_laast = 1 WHERE faknr = " & faknr & " AND shadowcopy = 1"
    oConn.execute(strSQL)
    end select

	
	if cint(momskonto) <> 2 then
	varKonto = intkontonr
	else
	varKonto = modkontonr
	end if
	
	
    intTotal = intBeloeb
    intMoms = intMoms
    intNetto = subtotaltilmoms
   
    
    oprid = 0

            '*** ******************** ***'
            '*** Opretter posteringer ***'
            '*** ******************** ***'
			
            
		    
            '**** Beregner Posteringer hvis faktura er godkendt *****'
		   '** Valuta / Kurs omregner til DKK **'
		    intNetto = (intNetto * (kurs/100))
		    intMoms = (intMoms * (kurs/100))
		    intTotal = (intTotal * (kurs/100))
            
            'Response.Write "intTotal: " & intTotal
            'Response.end

            '**** Postering debit konto ****'
            if intType = 0 then '(faktura)
            intNettoDeb = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
            intMomsDeb = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
            intTotalDeb = replace(replace(formatnumber(intTotal+intMoms, 2), ".", ""), ",", ".")
            else
            intNettoDeb = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
            intMomsDeb = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
            intTotalDeb = replace(replace(formatnumber(-(intTotal+intMoms), 2), ".", ""), ",", ".")
            end if 
            

            'Response.Write   intTotalDeb & "<br>"
            'Response.end
		    
            call opretPosteringSingle(oprid, "2", "dbopr", intkontonr, modkontonr, intTotalDeb, intTotalDeb, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)

		    
		    
		    
    		'**** Postering kredit konto ****'
    		'if intType = 0 then '(faktura)
    	'	intNettoKre = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
           ' 'intMomsKre = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
            'intTotalKre = replace(replace(formatnumber(-(intTotal+intMoms), 2), ".", ""), ",", ".")
            'else
            'intNettoKre = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
            'intMomsKre = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
            'intTotalKre = replace(replace(formatnumber(intTotal+intMoms, 2), ".", ""), ",", ".")
            'end if

            if intType = 0 then '(faktura)
            'intNettoKre = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
            intNettoKre = replace(replace(formatnumber(-intTotal), ".", ""), ",", ".")
		    intMomsKre = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
		    intTotalKre = replace(replace(formatnumber((-intTotal/1 - intMoms/1)), ".", ""), ",", ".")
                                                
            else
            'intNettoKre = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
            intNettoKre = replace(replace(formatnumber(intTotal), ".", ""), ",", ".")
		    intMomsKre = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
		    intTotalKre = replace(replace(formatnumber((intTotal/1 - intMoms/1)), ".", ""), ",", ".")
                                                
            end if
            
            
            'Response.Write   intTotalKre
            'Response.end

            call opretPosteringSingle(oprid, "2", "dbopr", modkontonr, intkontonr, intNettoKre, intNettoKre, intMomsKre, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)
    		
    		'*** Opretter posteringer på Execon / Immenso version ***'
		    vorrefId = vorref
		    intBeloeb = intTotalDeb
		    call ltoPosteringer
							
		    'Response.end

		     
	'***** Slut konti, moms og posteringer ***'
	'*****************************************'
	
	
    '*** Sætter opr ID fra posteringer ***'
    '*** brugs hvis fortryd godkend benyttes, til at slette evt. posteringer ****'
    
    strSQL = "UPDATE fakturaer SET oprid = "& oprid &" WHERE fid = " & fakid
    oConn.execute(strSQL)


    '**** Opdaterer faktura nr rækkefølge *****'
    call opdater_fakturanr_rakkefgl(opdFaknrSerie, intFaknumFindes, sqlfld, intFaknum)

	
    
   
    Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
    'Response.Write("<script language=""JavaScript"">window.opener.top.frames['erp2_2'].location.reload();</script>")
    
    select case lto
    case "execon", "immenso"', "intranet - local"
    Response.redirect "erp_make_pdf_multi.asp?fakids="&fakid&"&makepdf=1&lto="&lto
    case else   
    Response.Write("<script language=""JavaScript"">window.close();</script>")
    end select
    

end if
%>

<!--include file="../inc/regular/footer_inc.asp"-->
