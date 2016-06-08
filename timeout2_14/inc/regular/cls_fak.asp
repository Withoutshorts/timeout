<% 


sub erpfortryd



    oprid = 0
	thisfakid = id
	
	strSQL = "SELECT oprid FROM fakturaer WHERE fid =" & id
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	oprid = oRec("oprid")
	oRec.movenext
	wend
	oRec.close
	
	if oprid <> 0 then
	strSQL = "DELETE FROM posteringer WHERE oprid = "& oprid
	oConn.execute(strSQL)
	end if
	
	
	strSQL = "UPDATE fakturaer SET betalt = 0 WHERE fid =" & id
	oConn.execute(strSQL)


       Response.redirect "erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&visfaktura=0"
	

end sub



sub erpslet

                slttxtalt = ""
	            slturlalt = ""
	
	            slttxt = "<br>Du er ved at <b>Slette</b> en <b>faktura-skrivelse!</b><br> Er du sikker på du ønsker at slette fakturaen? Fakturaen kan ikke genskabes."
	
                if rdir = "hist" then
                slturl = "erp_opr_faktura_fs.asp?func=sletja&visjobogaftaler=0&visfaktura=0&visminihistorik=0&nomenu=1&id="&id&"&rdir="&rdir 
                else
	            slturl = "erp_opr_faktura_fs.asp?func=sletja&visjobogaftaler=1&visminihistorik=1&id="&id&"&rdir="&rdir
	            end if
               

	            call sltque_small(slturl,slttxt,slturlalt,slttxtalt,sletLft,sletTop)



end sub


sub erpsletja

     '** Sletter SHADOW COPY ****
	                strSQLshadow = "SELECT faknr, kommentar, jobid FROM fakturaer WHERE fid = "& id
                    oRec.open strSQLshadow, oConn, 3
                    if not oRec.EOF then
                    intFaknum = oRec("faknr")
                    jobidSlet = oRec("jobid")
            
                            '*** Indsætter i delete historik ****'
	                        call insertDelhist("fak", id, intFaknum &" (jobid: " &jobidSlet &")", "xx", session("mid"), session("user"))
    
                    end if
                    oRec.close
    
                     strSQLshadowCopy = "DELETE FROM fakturaer WHERE faknr = "& intFaknum & " AND shadowcopy = 1" 
                    oConn.execute(strSQLshadowCopy)
                    '***'
    
	
	                '*** Evt. Posteringer slettes under "fortryd godkend"
	
	                '*** Sletter Faktura KUN KLADER PGA bogføringslov ************
	                oConn.execute("DELETE FROM fakturaer WHERE Fid = "& id &"")
	
                
	
	                oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
	                oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& id &"")
	
                    '** renser db, opdater materiale forbrug ***'
                    strSQLselmf = "SELECT matfrb_id FROM fak_mat_spec WHERE matfakid = " & id
                    oRec4.open strSQLselmf, oConn, 3
                    while not oRec4.EOF
	
                    '*** Markerer i materialeforbrug at materiale er faktureret ***'
                    strSQLmf = "UPDATE materiale_forbrug SET erfak = 0 WHERE id = " & oRec4("matfrb_id")
                    oConn.execute(strSQLmf)
	
                    oRec4.movenext
                    wend
                    oRec4.close
	
	
                    '*** Sletter hidtidige regs ***'
                    strSQLdel = "DELETE FROM fak_mat_spec WHERE matfakid = " & id
                    oConn.execute strSQLdel
	
	
                    'response.write "erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&visfaktura=0"
                    'response.end

                   'response.write "rdir: "& rdir
                    'response.end            

                   if rdir = "hist" then
                   Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
                   Response.Write("<script language=""JavaScript"">window.close();</script>")
                   else
	               Response.redirect "erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&visfaktura=0"
                   end if        

end sub




public intFaknum, intFaknumFindes, opdFaknrSerie, sqlfld
function findFaknr(func)

    

    strSQL = "SELECT fakturanr, kreditnr, fakturanr_kladde FROM licens WHERE id = 1"
	        oRec.open strSQL, oConn, 3
            if not oRec.EOF then

	                if func = "dbopr" then
        	        
        	                intFaknumFindes = 0
	                        intFaknum = 0
                            opdFaknrSerie = 1 
        	        
        	            
	                   
                                        
                                                if cint(intFakbetalt) = 0 OR cint(medregnikkeioms) = 1 OR cint(medregnikkeioms) = 2 then ' kladde / intern / Handelsfak
                                    
                                                     intFaknum = oRec("fakturanr_kladde") + 1
	                                                 sqlfld = "fakturanr_kladde"

                                                else '*** Godkendt (ikke intern)

                                                    select case intType
	                                                case 0
	                                                intFaknum = oRec("fakturanr") + 1
	                                                sqlfld = "fakturanr"
	                                                case 1
        	            
	                                                if oRec("kreditnr") <> "-1" then '*** Samme nummer rækkkefølge **'
	                                                intFaknum = oRec("kreditnr") + 1
	                                                sqlfld = "kreditnr"
	                                                else
	                                                intFaknum = oRec("fakturanr") + 1
	                                                sqlfld = "fakturanr"
	                                                end if
        	            
	                                                end select
                        
                                        
                                                end if
                                
                                
                                       
                        
                  
                    
                  
                    
                    else 'rediger
                
                        '** Hvis status skift fra kladde til godkendt (en godkendt kan ikke redigeres)
                        '** Hvis skift fra Intern til ikke intern (kladde eller godkendt)
            
                    opdFaknrSerie = 0
                    intFaknumFindes = 0

                    if func = "dbred_gk" then
                    intFaknum = faknr
                    else
                        fak_laast = request("FM_fak_laast")
                        intFaknum = request("faknr")
                    end if
                 
                                         
                                         if cint(intFakbetalt) = 1 AND cint(fak_laast) = 0 then '** GodKendt og har faktura været godkendt engang må den ikke skifte faknr **'


                                            if (cint(medregnikkeioms) = 1 OR cint(medregnikkeioms) = 2) AND cint(medregnikkeioms_opr) = 0 then ' --> nu intern 

                                                intFaknum = oRec("fakturanr_kladde") + 1
	                                            sqlfld = "fakturanr_kladde" 

                                                opdFaknrSerie = 1
                                                

                                            else
                                                    
                                                    

                                                    if (cint(medregnikkeioms) <> 1 AND cint(medregnikkeioms) <> 2) then ' intern (behold intern nummer selvom godkendt)

                                                    'Response.Write "her: " & intType
                                                   
                                                    'Nu godkendt og IKKE Intern
                                                    select case intType
	                                                case 0
	                                                intFaknum = oRec("fakturanr") + 1
	                                                sqlfld = "fakturanr"
	                                                case 1
        	            
	                                                if oRec("kreditnr") <> "-1" then '*** Samme nummer rækkkefølge **'
	                                                intFaknum = oRec("kreditnr") + 1
	                                                sqlfld = "kreditnr"
	                                                else
	                                                intFaknum = oRec("fakturanr") + 1
	                                                sqlfld = "fakturanr"
	                                                end if
        	            
	                                                end select

                                                    opdFaknrSerie = 1

                                                    end if

                                                    'Response.Write "intFaknum: "& intFaknum
                                                    'Response.end

                                            end if

                                             

                                         end if ' GK

                                  

                                    
                            
                            
                    
                  end if' **rediger/opr
            
            end if
            oRec.close

           'Response.flush

                        '*************************************
                        '*** Opdaterer rækkefølge ***********'
                        '*************************************

                        ''*** Må fakruaer have samme rækkefølge ***'
                     

                           strSQL = "SELECT faknr FROM fakturaer WHERE faknr = '"& intFaknum &"' AND fid <> "& id & " AND shadowcopy = 0"
                            'Response.Write strSQL
                            'Response.flush
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF
                            intFaknumFindes = 1
                            oRec.movenext
                            wend
                            oRec.close

                            
                        
                
                        'Hvis NT og faktura oprettelse er en rediger faktura (fungerer som opret ny)       
                        'if (lto = "intranet - local" OR lto = "nt") AND len(trim(request("faknr"))) <> 0 then

                        
                               
                        '       intFaknum = request("faknr")

                        'else

                            


                        '**** Opdaterer fakturanr i licens tabel ****'
					    'if cint(opdFaknrSerie) = 1 AND cint(intFaknumFindes) = 0 then
                        'strSQL = "UPDATE licens SET "& sqlfld &" = "& intFaknum &" WHERE id = 1"
                        'oConn.execute(strSQL)
                        'end if

                        'end if
                        

            '**** Faktura nr SLUT **'

            


end function

'**** Opdaterer faktura nr rækkefølge *****'
function opdater_fakturanr_rakkefgl(opdFaknrSerie, intFaknumFindes, sqlfld, intFaknum)

    if cint(opdFaknrSerie) = 1 AND cint(intFaknumFindes) = 0 then
    strSQL = "UPDATE licens SET "& sqlfld &" = "& intFaknum &" WHERE id = 1"
    oConn.execute(strSQL)
    end if


end function


public faktureret, faktureretKre, faktureretTimerEnhStk, faktureretLastFakDato
function stat_faktureret_job(jobid, sqlDatostart, sqlDatoslut)


  '*** Faktureret **'
		faktureret = 0
		faktureretTimerEnhStk = 0
        faktureretLastFakDato = day(sqlDatostart) & "/"& month(sqlDatostart) &"/"& year(sqlDatostart) '"2002-01-01"
		
		'*** Faktureret ***'
		strSQL2 = "SELECT if(faktype = 0, f.beloeb * (f.kurs/100), f.beloeb * -1 * (f.kurs/100)) AS faktureret, if(faktype = 0, SUM(fd.aktpris * (fd.kurs/100)), SUM(fd.aktpris * -1 * (fd.kurs/100))) AS aktbel, fakdato " _
		&" FROM fakturaer AS f "_
        & "LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"_
		&" WHERE jobid = "& jobid &" AND aftaleid = 0 AND shadowcopy = 0"
		
		
		if realfakpertot <> 0 then
		strSQL2 = strSQL2 &" AND ((brugfakdatolabel = 0 AND fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"')"
        strSQL2 = strSQL2 &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'))"
		end if
		
      
        strSQL2 = strSQL2 &" GROUP BY fid ORDER BY fakdato"

        'Response.Write strSQL2
        'Response.end
		oRec2.open strSQL2, oConn, 3
		
		while not oRec2.EOF
		faktureret = faktureret + oRec2("faktureret")
        

        if cDate(oRec2("fakdato")) < cDate("01-06-2010") AND lto = "epi" then
        faktureretTimerEnhStk = faktureretTimerEnhStk + oRec2("faktureret")
        else
        faktureretTimerEnhStk = faktureretTimerEnhStk + oRec2("aktbel")
        end if

        '*** Bruger altid system dato til beregneing af igangværende arbejde, da timer kan indtastes på job fra systemlukkedato = fakturadato
        faktureretLastFakDato = oRec2("fakdato")
        
	    oRec2.movenext
        wend
		oRec2.close
		



		'*** Kredit ***'
		'*** Beregnes ovenfor ***'
		
        'strSQL2 = "SELECT f.beloeb * (f.kurs/100) AS faktureret, SUM(fd.aktpris * (fd.kurs/100)) AS aktbel, fakdato " _
		'&" FROM fakturaer AS f "_
        '& "LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"_
		'&" WHERE jobid = "& jobid &" AND faktype = 1 AND aftaleid = 0 AND shadowcopy = 0"
		
		
		
		

        'if realfakpertot <> 0 then
        'strSQL2 = strSQL2 &" AND fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
       'end if
		
        'faktureretKre = 0
        'KreTimerEnhStk = 0

		'strSQL2 = strSQL2 &" GROUP BY fid"
		
        'oRec2.open strSQL2, oConn, 3
		
        'while not oRec2.EOF
		'faktureretKre = faktureretKre + oRec2("faktureret")
            
         '   if oRec2("fakdato") < "01-06-2010" AND lto = "epi" then
         '   KreTimerEnhStk = KreTimerEnhStk + oRec2("faktureret")
         '   else
         '   KreTimerEnhStk = KreTimerEnhStk + oRec2("aktbel")
         '   end if

	    'oRec2.movenext
        'wend
		'oRec2.close
		
		'if faktureret <> 0 then
		'faktureret = faktureret - (faktureretKre)
		'else
		'faktureret = 0
		'end if


        if faktureretTimerEnhStk <> 0 then
        faktureretTimerEnhStk = faktureretTimerEnhStk - (KreTimerEnhStk)
        else
        faktureretTimerEnhStk = 0
        end if
        

end function
%>