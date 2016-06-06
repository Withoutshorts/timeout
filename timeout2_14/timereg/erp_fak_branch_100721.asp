<%response.buffer = true%>

<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<!--#include file="../inc/xml/erp_fak_xml_inc.asp"--> 





<%

'Response.write lto

level = session("rettigheder")


if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FN_getKpers"
                             if len(trim(request.Form("cust"))) <> 0 then
                             kid = request.Form("cust")
                             else
                             kid = 0
                             end if
                             
                             if len(request("att_val")) <> 0 then
                             strAtt = request("att_val")
                             else
                             strAtt = 0
                             end if
                             
                                strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& kid &" ORDER BY navn"
				                'Response.write strSQLkpersCount
				                'Response.end
				                oRec2.open strSQLkpersCount, oConn, 3
                					
					                while not oRec2.EOF
						                if cint(strAtt) = cint(oRec2("id")) then
						                attSel = "SELECTED"
						                else
						                attSel = ""
						                end if
						                
						                '*** ÆØÅ **'
						                kpersnavn = oRec2("navn")
						                call jq_format(kpersnavn)
                                        kpersnavn = jq_formatTxt 
						             
						                
						                
						                Response.Write "<option value='"& oRec2("id") &"' "& attSel &">"& kpersnavn &"</option>"
						                
					                oRec2.movenext
					                wend
					                oRec2.close
					                
					                %>
					                <%if strAtt = "991" then
				                selth1 = "SELECTED"
				                else
				                selth1 = ""
				                end if%>	
				                <option value="991" <%=selth1%>>Den &oslash;konomi ansvarlige</option>
				                <%if strAtt = "992" then
				                selth2 = "SELECTED"
				                else
				                selth2 = ""
				                end if%>	
				                <option value="992" <%=selth2%>>Regnskabs afd.</option>
				                <%if strAtt = "993" then
				                selth3 = "SELECTED"
				                else
				                selth3 = ""
				                end if%>	
				                <option value="993" <%=selth3%>>Administrationen</option>
					                
					           
                                    <%
                             
case "FN_getMatTilFak"
                            
                            lto = request("jq_lto")
                            func = request("jq_func")
                            stdatoKri = request("jq_stdato")
                            slutdato = request("jq_sldato")
                            jobid = request("jq_jobid")
                            id = request("jq_id")
                            kuneks = request("jq_kuneks")
                            ingper = request("jq_ignper")
                            jq_intType = request("jq_inttype")
                            
                            valutaKursFak = request("jq_valutaValkurs")
                            valutaLabelFak = request("jq_valutafaklabel")
                            'Response.Write func
                            'Response.end
                            matSubTotalAlluMoms = 0
                            matSubTotalAll = 0
                           
                            %>                            
                            <table width=100% cellspacing=0 cellpadding=2 border=0>
	                        </tr>
	                        <tr bgcolor="orange">
	                        <%call matFelter
                            
                             '** Allerede fakturarede materialer 
	                        m = 0
	                        isMatWrt = " matid <> -1 "
	                        lastgrpnavn = ""
                        	
	                        if func = "red" then
                        	
	                        
	                        strSQL = "SELECT fms.matid, matnavn, matantal, matenhedspris AS matsalgspris, "_
	                        &" matenhed, matvarenr, ikkemoms, fms.valuta AS valutaid, fms.kurs, v.valutakode, matrabat, "_
	                        &" matfrb_mid AS mfusrid, matfrb_id AS mfid, matgrp, mgp.navn AS matgrpnavn, matbeloeb FROM fak_mat_spec fms "_
	                        &" LEFT JOIN valutaer v ON (v.id = fms.valuta) "_
	                        &" LEFT JOIN materiale_grp AS mgp ON (mgp.id = fms.matgrp) "_
	                        &" WHERE matfakid = "& id &" ORDER BY mgp.navn, matnavn"
	                        'Response.Write strSQL
	                        'Response.flush
	                        oRec.open strSQL, oConn, 3
                            while not oRec.EOF
                            
                            intMatRabat = (100*oRec("matrabat"))
                            valutaMatId = oRec("valutaid")
                            valutaMatKode = oRec("valutakode") 
                            valutakodeSEL = valutaLabelFak 'valutaMatKode
                            valutaMatKurs = oRec("kurs")
                            ikkemoms = oRec("ikkemoms")
                            
                             '*** ÆØÅ **'
                             matnavn = oRec("matnavn")
                             call jq_format(matnavn)
                             matnavn = jq_formatTxt 
                            
                               
                            
                            
                            call materialer(1)
                            
                            isMatWrt = isMatWrt & " AND matid <> " & oRec("mfid")
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
                        	
                        	
	                        
                            end if
                            
                            
                            '** Ikke fakturerede Materialer / Eller v. ny faktura = alle 
	                        strSQL = "SELECT mf.matid, mf.matnavn, sum(mf.matantal) AS matantal, mf.matsalgspris, mf.matenhed, "_
	                        &" mf.matvarenr, mf.valuta, v.valutakode, v.kurs, mf.erfak, mf.id AS mfid, "_
	                        &" mf.usrid AS mfusrid, intkode, matgrp, mgp.navn AS matgrpnavn "_
	                        &" FROM materiale_forbrug mf "_
	                        &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	                        &" LEFT JOIN materiale_grp AS mgp ON (mgp.id = matgrp) "_ 
	                        &" WHERE mf.jobid = "& jobid 
	                        
	                        if kuneks <> 0 then 'kun eksterne
	                        strSQL = strSQL &" AND mf.intkode = 2 " 
	                        end if
	                        
	                        if ingper <> 1 then
	                        strSQL = strSQL &" AND forbrugsdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"'"
	                        end if
	                        
	                        '***hent kun dem der ikke allerede er faktureret ***'
	                        strSQL = strSQL &" AND erfak = 0 " 
	                        
	                        strSQL = strSQL &" AND ("& isMatWrt &") GROUP BY mf.matid, mf.matsalgspris, mf.matnavn ORDER BY mgp.navn, mf.matnavn"
	                        
	                        'Response.Write strSQL
	                        'Response.flush
	                        
	                        oRec.open strSQL, oConn, 3
                            im = 0
                            while not oRec.EOF
                            
                            
                            '*** ÆØÅ **'
                            matnavn = oRec("matnavn")
                            call jq_format(matnavn)
                            matnavn = jq_formatTxt 
                            
                            
                                
            
                         
                               
                            intMatRabat = intRabat
                            valutaMatId = oRec("valuta") 
                            valutaMatKode = oRec("valutakode") 
                            valutakodeSEL = valutaLabelFak 'valutaMatKode
                            valutaMatKurs = oRec("kurs")
                            
                            if (lto = "execon" OR lto = "immenso") then
                            ikkemoms = 1
                            else
                            ikkemoms = 0    
                            end if
                            
                            call materialer(2)
                            
                            
                            
                            im = im + 1
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
                           
                            if m = 0 then%>
	                        <tr><td colspan=9><br />(Der er ikke fundet nogen materialer i det valgte interval.)</td></tr>
	                        <%end if  %>
                                
                             <input id="jq_matbeltot" value="<%=matSubTotalAll %>" type="hidden" />
                        	 <input id="jq_matbeltot_umoms" value="<%=matSubTotalAlluMoms %>" type="hidden" />
                        	 
                             <input id="FM_antal_materialer_ialt" name="FM_antal_materialer_ialt" value="<%=m%>" type="hidden" />
	                        </table>
                            <%
                            

case "FN_getCustDesc"
            
            if request("vis_cvr") = "1" then
            vis_cvr = 1
            else
            vis_cvr = 0
            end if
            
            if request("vis_land") = "1" then
            vis_land = 1
            else
            vis_land = 0
            end if
            
            
            '*** Vælg kontakt eller filial / kpers
            if Request("attkid") = "kid" then
          
                
                 if len(trim(request.Form("cust"))) <> 0 then
                 kid = request.Form("cust")
                 else
                 kid = 0
                 end if
                 
                  strSQLkpersAdr = "SELECT kid, kkundenavn AS navn, adresse, postnr, city AS town, "_
                  &" land, ean, cvr FROM kunder WHERE kid = "& kid
		   
            else
                 
                 if len(trim(request("att_val"))) <> 0 then
                 att = request("att_val")
                 else
                 att = 0
                 end if
          
            strSQLkpersAdr = "SELECT kp.id, kp.navn, kp.adresse, kp.postnr, kp.town, kp.land, k.ean, k.cvr "_
            &" FROM kontaktpers AS kp "_
            &" LEFT JOIN kunder AS k ON (k.kid = kp.kundeid) WHERE kp.id = "& att
		   
            end if
            
            
            'Response.write strSQLkpersAdr
            'Response.end
             
            
            oRec.open strSQLkpersAdr, oConn, 3
            if not oRec.EOF then
                            
                            
                            if cint(vis_land) = 1 AND len(trim(oRec("land"))) <> 0 then
			                land = trim("<br>"&oRec("land"))
			                else
			                land = ""
			                end if
                            
                            
                            if len(trim(oRec("ean"))) <> 0 then
			                ean = trim("<br>EAN: "&oRec("ean"))
			                else
			                ean = ""
			                end if
			                
			                if cint(vis_cvr) = 1 AND len(trim(oRec("cvr"))) <> 0 then
			                cvr = trim("<br>CVR: "&oRec("cvr"))
			                else
			                cvr = ""
			                end if
			               
            
            strAdr = oRec("navn") & "<br>"& oRec("adresse") & "<br>"& oRec("postnr") & " " &oRec("town") & land & ean & cvr
            end if
            oRec.close
            
             '*** ÆØÅ **'
            call jq_format(strAdr)
            strAdr = jq_formatTxt
            
           Response.Write strAdr
            
End select
Response.End
end if


dim content
a = 0
dim aval
redim aval(a)

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	if len(request("FM_type")) <> 0 then
	intType = request("FM_type")
	else
	intType = 0
	end if
	
	
	
    thisfile = "erp_fak.asp"
    print = request("print")
    err = 0
    
    kid = request("FM_kunde")
   
    if len(trim(request("FM_job"))) <> 0 then
    jobid = request("FM_job")
    else
    jobid = 0
    end if
    
    if len(trim(request("FM_aftale"))) <> 0 then
    aftid = request("FM_aftale")
    else
    aftid = 0
    end if
    
    'jobelaft = request("jobelaft")
	
	func = request("func")
	
	
	call setFakPreInd()
	
	
	if len(trim(request("rdir"))) <> 0 then
	rdir = request("rdir")
	else
	rdir = ""
	end if
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%
	
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>Slette - (nulstille)</b> en <b>faktura-skrivelse</b>. Er dette korrekt?<br><br>Alle beløb på fakturaen vil blive nulstillet, og fakturaen vil figurere som en NUL-faktura. Faktura dato vil blive sat til <b>1. jan. 2002.</b>&nbsp;"
	
	slturl = "erp_fak.asp?FM_aftale="&aftid&"&FM_kunde="&kid&"&func=sletok&id="&id&"&FM_job="&jobid&"&FM_usedatointerval=1&FM_start_dag="&request("FM_start_dag")&"&FM_start_mrd="&request("FM_start_mrd")&"&FM_start_aar="&request("FM_start_aar")&"&FM_slut_dag="&request("FM_slut_dag")&"&FM_slut_mrd="&request("FM_slut_mrd")&"&FM_slut_aar="&request("FM_slut_aar")&"&rdir="&rdir
	
	call sltque_small(slturl,slttxt,slturlalt,slttxtalt,10,10)
	
	
	case "sletok"
	
	
	
	
	'** Sletter SHADOW COPY ****
	strSQLshadow = "SELECT faknr, kommentar FROM fakturaer WHERE fid = "& id
    oRec.open strSQLshadow, oConn, 3
    if not oRec.EOF then
    intFaknum = oRec("faknr")
            
            '*** Indsætter i delete historik ****'
	        call insertDelhist("fak", id, intFaknum, left(oRec("kommentar"), 50), session("mid"), session("user"))
    
    end if
    oRec.close
    
     strSQLshadowCopy = "DELETE FROM fakturaer WHERE faknr = "& intFaknum & " AND shadowcopy = 1" 
    oConn.execute(strSQLshadowCopy)
    '***'
    
	
	'*** Evt. Posteringer slettes under "fortryd godkend"
	
	'*** Sletter NULSTRILLER Faktura da den ikke må slettes i forhold til bogførings lov. ************
	'oConn.execute("DELETE FROM fakturaer WHERE Fid = "& id &"")
	oConn.execute("UPDATE fakturaer SET timer = 0, beloeb = 0, fakdato = '2002-01-01', Kommentar = 'Nulstillet', jobbesk = 'Nulstillet', konto = 0, modkonto = 0, moms = 0, timersubtotal = 0, "_
	&" subtotaltilmoms = 0, valuta = 1, kurs = 100, erfakbetalt = 0  WHERE Fid = "& id &"")
	
	
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
	
	
	if rdir = "hist" then
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	else
	%>
    <script>
	//alert("her")
	//window.top.frames['erp4_2'].location.href = "erp_opr_faktura_blank.asp"
	window.top.frames['erp3'].location.href = "erp_opr_faktura_blank2.asp"
    window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>"
    </script>
	<%
	end if
	
	
	
	
	case "fortryd"
	
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
	
	
	
	%>
	<script>
	window.top.frames['erp3'].location.href = "erp_fak.asp?func=red&id=<%=thisfakid%>&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>"
    window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?id=<%=thisfakid%>&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>"
   </script>
		
	<%
	
	
	
	
	case "dbred", "dbopr" 
	
				
	function SQLBlessDOT(s)
	dim tmpdot
	tmpdot = s
	tmpdot = replace(tmpdot, ".", "")
	SQLBlessDOT = tmpdot
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
	
	function SQLBlessPunkt(s)
	dim tmpPunkt
	tmpPunkt = s
	tmpPunkt = replace(tmpPunkt, ".", ",")
	SQLBlessPunkt = tmpPunkt
	end function
	
	function SQLBlessK(s)
    dim tmp
    tmp = s
    tmp = replace(tmp, "'", "''")
    SQLBlessK = tmp
    end function

    '************************
    '*** findes allerede ****
    '************************
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
	
	intBeloeb = replace(Request("FM_Beloeb"), ",", ".")
	
	'** Bruges denne ?? ***'
	if len(trim(Request("FM_Beloeb_umoms"))) <> 0 then
	intBeloebUmoms = replace(Request("FM_Beloeb_umoms"), ",", ".")
	else
	intBeloebUmoms = 0
	end if
	
	intTimer = SQLBless(Request("FM_Timer"))
	
	
	if aftid <> 0 then
	intRabat = Request("FM_rabat")
	else
	intRabat = 0
	end if
	
	'valutaogkurs = split(request("FM_valuta"),":")
	
	'for i = 0 TO UBOUND(valutaogkurs)
	'    if i = 0 then
	'    valuta = valutaogkurs(i)
	'    else
	'    kurs = replace(valutaogkurs(i), ",", ".")
	'    end if
	'next
	
	valuta = request("FM_valuta_all_1")
	kurs = replace(request("valutakurs_"& valuta &""), ",", ".")
	
	
    if request("FM_sprog") <> 1 then
	intSprog = request("FM_sprog")
	else
	intSprog = 1
	end if
	
	
	
	        '******************'
	        '*** Faktura nr ***'
	        '******************'
        	
	        if func <> "dbred" then
        	        
        	        intFaknumFindes = 0
	                intFaknum = 0 
        	        
        	            
	                    strSQL = "SELECT fakturanr, kreditnr FROM licens WHERE id = 1"
	                    oRec.open strSQL, oConn, 3
                        while not oRec.EOF
                        
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
                        
                        
                        oRec.movenext
                        wend
                        oRec.close
                        
                        strSQL = "SELECT faknr FROM fakturaer WHERE faknr = '"& intFaknum &"'" 
                        'Response.Write strSQL
                        'Response.flush
                        oRec.open strSQL, oConn, 3
                        while not oRec.EOF
                        intFaknumFindes = 1
                        oRec.movenext
                        wend
                        oRec.close
                        
                  
                    
                  
                    
            else
            
            intFaknumFindes = 0
            intFaknum = request("faknr")
                   
                    
            end if
            
            '**** Faktura nr SLUT **'
	
	
	
	useleftdiv = "m"
	
	'*** Her tjekkes om alle required felter er udfyldt. ***
	if len(Request("FM_timer")) = 0 OR len(Request("FM_beloeb")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

	<%
	'len(Request("FM_faknr")) = 0 OR
	errortype = 26
	call showError(errortype)
	
	else%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(trim(intBeloeb))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
			<%
			errortype = 27
			call showError(errortype)
			
			isInt = 0
					else
					
					call erDetInt(trim(intTimer))
					if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					
					<%
					errortype = 28
					call showError(errortype)
					isInt = 0
							
							else
								
								
						     if intFaknumFindes = 1 then
						     
						     %>
					        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					        
					        <%
					        errortype = 89
					        call showError(errortype)
					        
							
							else
							
							
							'*** Hvis alle required er udfyldt ***'
						            
									'**************************************'
						            '** Faktura dato, Periode og Label ****'
						            '**************************************'
						            
						            
							        istDato = request("FM_istdato_aar") & "/" & request("FM_istdato_mrd") & "/" &request("FM_istdato_dag") 
									
									showistDato = request("FM_istdato_dag") & "/" & request("FM_istdato_mrd") & "/" &request("FM_istdato_aar") 
									
									istDato2 = request("FM_istdato2_aar") & "/" & request("FM_istdato2_mrd") & "/" &request("FM_istdato2_dag") 
									showistDato2 = request("FM_istdato2_dag") & "/" & request("FM_istdato2_mrd") & "/" &request("FM_istdato2_aar") 
									
									labelDato = request("FM_labelDato_aar") & "/" & request("FM_labelDato_mrd") & "/" &request("FM_labelDato_dag") 
									
									
									fakDato = Request("FM_start_aar") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_dag") '& time 
									
									if len(request("FM_brugfakdatolabel")) <> 0 then 
							        brugfakdatolabel = 1
							        else
							        brugfakdatolabel = 0
							        end if
									
									'*** Brug label eller faktura system dato **'
									if cint(brugfakdatolabel) = 1 then
									showfakDato = labelDato 'showistDato2
									else
									showfakDato = Request("FM_start_dag") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_aar")
									end if
		                            
									betbetint = Request.Form("FM_betbetint")
		                            
									
									
									
									'Response.Cookies("erp")("forfaldsdatoDage") = adddayVal
									
									'*** Sidste rettidige bet. dato *'
									if betbetint <> 0 then
									reformatdate = split(Request("FM_forfaldsdato"), ".")
									forfaldsdato = reformatdate(0) & "/" & reformatdate(1) & "/" & reformatdate(2)
									else
									forfaldsdato = DateAdd("d", -1, showfakDato)
									end if
									dtb_dato = year(forfaldsdato) & "/"& month(forfaldsdato) & "/" & day(forfaldsdato)
									
									strEditor = session("user")
									strDato = session("dato")
									intFakbetalt = request("FM_betalt")
									intfakadr = request("FM_Kid")
									
									kundeid = kid 'request("FM_kundeid")
									
									if request("FM_att_vis") = "on" then
									strAtt = request("FM_att")
									else
									strAtt = 0
									end if
									
									if intFakbetalt <> 1 then
									intFakbetalt = 0
									else
									intFakbetalt = intFakbetalt 
									end if
									
									if len(request("FM_varenr")) <> 0 then
									strVarenr = request("FM_varenr")
									else
									strVarenr = 0
									end if
									
									intEnhedsang = request("FM_enheds_ang")
																		
									jobfaktype = request("jobfaktype")
									
									
									
									'************************************'
									'***** Konti, moms og posteringer ***'
									'************************************'
									
						
									'*** Debit og Kredit konto **'
									if len(request("FM_kundekonto")) <> 0 then
									varKonto = request("FM_kundekonto")
									else
									varKonto = 0
									end if
									
									'Response.Write "varKonto "& varKonto
									'Response.end
									
									Response.Cookies("erp")("debkonto") = varKonto
									
									if len(request("FM_modkonto")) <> 0 then
									varModkonto = request("FM_modkonto")
									else
									varModkonto = 0
									end if
									Response.Cookies("erp")("krekonto") = varModKonto
									
									'** Momskonto **'
									'*** Hvilken konto bruges til at beregne moms på fakturaen? ***'
									'if len(request("FM_momskonto")) <> 0 then
	                                'momskonto = request("FM_momskonto")
	                                'else
	                                'momskonto = "1"
	                                'end if
                                	
	                                'Response.Cookies("erp")("momskonto") = momskonto
	                                
									
									'*** Data til postering ***'
						            if aftid <> 0 then
						            posteringsTxt = left(request("FM_jobbesk"), 20)
						            else
						            posteringsTxt = left(request("FM_jobnavn"), 20)
						            end if
							
							      
							        strTekst = posteringsTxt
							        vorrefId = session("Mid")
							        intStatus = 1
							        
							       
							        intkontonr = varKonto
									modkontonr = varModKonto
									
									momssats = request("FM_momssats")
									
									'if momskonto = "1" then
									momskontoUse = 0 'intkontonr
									momskonto = momskontoUse
									'else
									'momskontoUse = modkontonr
									'end if 
									
									
									'*** Til momsberegning ***'
									'*** d = Beregner moms oven i et netto beløb.
									'*** k = Trækker moms fra et brutto beløb.
									'****'
									
									debkre = "fak"
								    
								    varBilag = intFaknum 
									posteringsdato = fakDato
									strKomm = SQLBlessK(Request("FM_Komm"))
									antalAkt = Request("antal_x")
									
									'subtotal = 0
									'intBeloeb = replace(intBeloeb, ".", ",")
									'Response.Write "intBeloeb " & intBeloeb &" intBeloebUmoms "& intBeloebUmoms
									'Response.end
									intTotalMoms = replace(intBeloeb, ".", ",") - (replace(intBeloebUmoms, ".", ","))
									
									'*** Momsberegning ****'
									momsBelob = (intTotalMoms * momssats/100)
									
									
									'** Momsberegning til faktura **'
									'call beregnmoms(debkre, intTotalMoms, momskontoUse)
								    
								    
								   
								    'intTotal = replace(intBelob, ".", ",")
								    intNetto = replace(intTotalMoms, ".", ",") 
	                                intMoms = momsBelob
	                                intTotal = replace(intBeloeb, ".", ",")
	                                intTotal = intTotal + intMoms
	                                
	                                'Response.Write "intMoms" & intMoms  & "<br>"
	                                'Response.Write "intTotalMoms" & intTotalMoms & "<br>"
	                                'Response.Write "intBeloeb" & intBeloeb & "<br>"
	                                'Response.Write  "intTotal" & intTotal & "<br>"
		                            
		                            'Response.end
		                            
		                            showmoms = replace(intMoms, ".", "")
								    subtotaltilmoms = replace(intTotalMoms, ",", ".")
								    
								  
								   
						                    '*** ******************** ***'
							                '*** Opretter posteringer ***'
							                '*** ******************** ***'
                							
							                if intFakbetalt <> 0 then 
                						    
						                    '**** Beregner Posteringer hvis faktura er godkendt *****'
                						    '** Valuta / Kurs omregner til DKK **'
                						    intNetto = (intNetto * (kurs/1)/100)
                						    intMoms = (intMoms * (kurs/1)/100)
                						    intTotal = (intTotal * (kurs/1)/100)
                						    
                						    'Response.Write "intTotal: " & intTotal
                                            'Response.end
                			
							                '**** Postering debit konto ****'
							                '**** Ændre komme til punktum til postering **'
							                if intType = 0 then '(faktura)
							                intNettoDeb = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
		                                    intMomsDeb = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
		                                    intTotalDeb = replace(replace(formatnumber(intTotal, 2), ".", ""), ",", ".")
                		                    else
                		                    intNettoDeb = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
		                                    intMomsDeb = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
		                                    intTotalDeb = replace(replace(formatnumber(-intTotal, 2), ".", ""), ",", ".")
                		                    end if 
							                
							                
							                
							                'Response.Write "intTotalDeb" & intTotalDeb
							                'Response.end
							                
                						    oprid = 0
                						    if intkontonr <> 0 then
		                                    
		                                        call opretPosteringSingle(oprid, "2", "dbopr", intkontonr, modkontonr, intTotalDeb, intTotalDeb, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
    		
                    						    
                    						    
                    						    
                    						    
                                    		    '**** Postering kredit konto ****'
                                    		    if intType = 0 then '(faktura)
                                    		    intNettoKre = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
		                                        intMomsKre = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
		                                        intTotalKre = replace(replace(formatnumber(-intTotal, 2), ".", ""), ",", ".")
                		                        else
                		                        intNettoKre = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
		                                        intMomsKre = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
		                                        intTotalKre = replace(replace(formatnumber(intTotal, 2), ".", ""), ",", ".")
                		                        end if
                    		                    
                    		                    
	                                            call opretPosteringSingle(oprid, "2", "dbopr", modkontonr, intkontonr, intNettoKre, intNettoKre, intMomsKre, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
                                        		
                                        		'Response.end
                                        		
                                    		end if '*** Kontonr <> 0
                						    
                						    
						                    end if
								            
								            
								            intMoms = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
								    
									'***** Slut konti, moms og posteringer ***'
									'*****************************************'
									
									'Response.end
									
									
							'******* Modtager *****'
							if request("FM_land_vis") = "on" then
							intModland = 1
							else
							intModland = 0
							end if
							
							if len(trim(request("FM_modtageradr"))) <> 0 then
							modtageradr = replace(request("FM_modtageradr"), "'","''")
							else
							modtageradr = ""
							end if
							
							if request("FM_att_vis") = "on" then
							intModAtt = 1
							else
							intModAtt = 0
							end if
							
							if len(trim(request("FM_usealtadr"))) <> 0 then
							usealtadr = request("FM_altadr")
							else
							usealtadr = 0
							end if
							
							'if request("FM_tlf_vis") = "on" then
							'intModTlf = 1
							'else
							intModTlf = 0
							'end if
							
							if request("FM_cvr_vis") = "on" then
							intModCvr = 1
							else
							intModCvr = 0
							end if
							
							
							'***** Afsender ****'
							if request("FM_yswift_vis") = "on" then
							intAfsSwift = 1
							else
							intAfsSwift = 0
							end if
							
							
							if request("FM_yiban_vis") = "on" then
							intAfsIban = 1
							else
							intAfsIban = 0
							end if
							
							if request("FM_ycvr_vis") = "on" then
							intAfsICVR = 1
							else
							intAfsICVR = 0
							end if
								
							if request("FM_yemail_vis") = "on" then
							intAfsEmail = 1
							else
							intAfsEmail = 0
							end if
							
							if request("FM_ytlf_vis") = "on" then
							intAfsTlf = 1
							else
							intAfsTlf = 0
							end if	
							
							if request("FM_yfax_vis") = "on" then
							intAfsFax = 1
							else
							intAfsFax = 0
							end if		
							
							if len(trim(request("FM_sideskiftlinier"))) <> 0 then
							sideskiftlinier = request("FM_sideskiftlinier")
							else
							sideskiftlinier = 0
							end if
							
							vorref = request("FM_vorref")
							
							'*** Vis kun totalbeløb ***'
							if len(request("FM_hidesumaktlinier")) <> 0 then
							hidesumaktlinier = request("FM_hidesumaktlinier")
							else
							hidesumaktlinier = 0
							end if
							
							
							'*** Jobbesk ****'
							if len(request("FM_visjobbesk")) <> 0 OR aftid <> 0 then
							jobBesk = SQLBlessK(request("FM_jobbesk"))
						    else
							jobBesk = ""
							end if
							
							'** vis IKKE jobnavn **'
							if len(trim(request("FM_visikkejobnavn"))) <> 0 then
							visikkejobnavn = 1
							else
							visikkejobnavn = 0
							end if
							
							
							'*** Cookie på hsu jobbesk ***'
							if request("FM_visjobbesk") = "1" AND func = "dbopr" then 
							response.Cookies("erp")("huskjobbesk") = "1"
							else
							response.Cookies("erp")("huskjobbesk") = "0"
							end if
							
							
							'*** Timer subtotal beløb ****
							if len(request("FM_timer_beloeb")) <> 0 then
							timersubtotal = SQLBless2(request("FM_timer_beloeb"))
							else
							timersubtotal = 0
							end if
							
							
							if len(request("FM_visjoblog")) <> 0 then
							visjoblog = 1
							else
							visjoblog = 0
							end if
							
							
							if len(request("FM_vis_joblog_timepris")) <> 0 then
							visjoblog_timepris = 1
							else
							visjoblog_timepris = 0
							end if
							
							if len(request("FM_vis_joblog_enheder")) <> 0 then
							visjoblog_enheder = 1
							else
							visjoblog_enheder = 0
							end if
							
							if len(request("FM_vis_joblog_mnavn")) <> 0 then
							visjoblog_mnavn = 1
							else
							visjoblog_mnavn = 0
							end if
							
							
							
							if len(request("FM_vismatlog")) <> 0 then
							vismatlog = 1
							else
							vismatlog = 0
							end if
							
							
							
							if len(request("FM_visrabatkol")) <> 0 then
							visrabatkol = 1
							else
							visrabatkol = 0
							end if
							
							if len(request("FM_visperiode")) <> 0 then 
							visperiode = 1
							else
							visperiode = 0
							end if
							
							
							if len(trim(request("FM_fak_ski"))) <> 0 then
							fak_ski = 1
							else
							fak_ski = 0
							end if
							
							if len(trim(request("FM_fak_abo"))) <> 0 then
							fak_abo = 1
							else
							fak_abo = 0
							end if
							
							if len(trim(request("FM_fak_ubv"))) <> 0 then
							fak_ubv = 1
							else
							fak_ubv = 0
							end if
							
							
							
							Response.Cookies("erp")("visperiode") = visperiode
							Response.Cookies("erp")("flabel") = visperiode
							
							'Response.Write "request(FM_visrabatkol)" & request("FM_visrabatkol")
							'Response.flush
							'Response.end
							
							
							'*** Materiale filter ****
							if request("FM_mat_viskuneks") <> "" then
	                        response.cookies("erp")("matvke") = request("FM_mat_viskuneks")
	                        else
	                        response.cookies("erp")("matvke") = ""
	                        end if
	                        
	                        
	                        if len(trim(request("FM_mat_ignper"))) <> 0 then
	                        response.cookies("erp")("matignper") = 1
	                        else
	                        response.cookies("erp")("matignper") = ""
	                        end if
	                       
							
							if len(trim(request("FM_showmatasgrp"))) <> 0 then
							showmatasgrp = 1
							else
							showmatasgrp = 0
							end if
									
							'************************************************************************************
							'*** Opdaterer / Redigerer faktura 													*
							'************************************************************************************
							if func = "dbred" then
							
											strSQLupd = "UPDATE fakturaer SET"_
										    &" fakdato = '"& fakDato &"', "_
											&" timer = "& intTimer  &", "_
											&" beloeb = "& intBeloeb &", "_
											&" kommentar = '"& strKomm &"', "_
											&" editor = '"& strEditor &"', "_
											&" dato = '"& strDato &"', "_
											&" tidspunkt = '23:59:59', "_
											&" betalt = "& intFakbetalt &", "_
											&" b_dato = '"& dtb_dato &"', "_
											&" fakadr = "& intfakadr &", "_
											&" att = '"& strAtt &"', "_
											&" konto = "& varKonto &", "_
											&" modkonto = "& varModkonto &", "_
											&" faktype = "& intType &", "_
											&" vismodland = "& intModland &", "_
											&" vismodatt = "& intModAtt &", "_
											&" vismodtlf = "& intModTlf &", "_
											&" vismodcvr = "& intModCvr &", "_
											&" visafstlf = "& intAfsTlf &", "_
											&" visafsemail = "& intAfsEmail &", "_
											&" visafsswift = "& intAfsSwift &", "_
											&" visafsiban = "& intAfsIban &", "_
											&" visafscvr = "& intAfsICVR &", "_
											&" moms = "& replace(showmoms,",",".") &", "_
											&" enhedsang = "& intEnhedsang &", "_
											&" varenr = '"& strVarenr &"', jobbesk = '"& jobBesk &"', "_
											&" timersubtotal = "& timersubtotal &", "_
											&" visjoblog = "& visjoblog &", visrabatkol = "& visrabatkol &", "_
											&" vismatlog = "& vismatlog &", rabat = "& intRabat &", "_
											&" visjoblog_timepris = "& visjoblog_timepris &", "_
											&" visjoblog_enheder = "& visjoblog_enheder &", "_
											&" visjoblog_mnavn = "& visjoblog_mnavn & ", "_
											&" visafsfax = "& intAfsFax &", "_
											&" subtotaltilmoms = "& subtotaltilmoms &", "_
											&" valuta = "& valuta &", kurs = "& kurs &", "_
											&" sprog = "& intSprog &", istdato = '"& istDato &"', "_
											&" momskonto = "& momskonto &", visperiode = "& visperiode &", "_
											&" jobfaktype = "& jobfaktype &", betbetint = "& betbetint &", "_
											&" brugfakdatolabel = "& brugfakdatolabel &", istdato2 = '"& istDato2 &"', "_
											&" momssats = "& momssats &", modtageradr = '"& modtageradr &"', "_
											&" usealtadr = "& usealtadr &", vorref = '"& vorref &"', fak_ski = "& fak_ski &", "_
											&" showmatasgrp = "& showmatasgrp &", "_
											&" hidesumaktlinier = "& hidesumaktlinier &", sideskiftlinier = "& sideskiftlinier &", "_
											&" labeldato = '"& labelDato &"', "_
											&" fak_abo = "& fak_abo &", fak_ubv = "& fak_ubv &", visikkejobnavn = "& visikkejobnavn &""_
											&" WHERE Fid = "& id 
											
											
											'Response.write (strSQLupd)
											'Response.end
											oConn.execute(strSQLupd)
											
											
												    '*** Shadowcopy **
													'*** Hvis der oprettes en faktura på en aftale, hvor der er tilknyttet job, skal der oprettes en shadowcopy 
													'*** På de job der er tilknyttet så de kan lukkes for indtastning, samt at man kan se at der ligger en GHOSTET
													'*** Faktura på jobbet 
													
													if aftid <> 0 then
													
													    
													    strSQLshadow = "SELECT faknr FROM fakturaer WHERE fid = "& id
													    oRec.open strSQLshadow, oConn, 3
                                                        if not oRec.EOF then
                                                        
                                                        intFaknum = oRec("faknr")
                                                        
                                                        end if
                                                        oRec.close
													    
													    strSQLshadowCopy = "UPDATE fakturaer SET fakdato = '"& fakDato &"', istdato = '"& istDato &"' WHERE faknr = "& intFaknum & " AND shadowcopy = 1" 
													    oConn.execute(strSQLshadowCopy)
													    
													
													    
													
													end if
											
											
							thisfakid = id
							end if
							'***
											
											
											
							'************************************************************************************
							'*** Opretter faktura *********														*
							'************************************************************************************
							if func = "dbopr" then
												
												
												'Response.Write "intType " & intType
												'Response.Flush
												
												if intType <> 0 then '*** kreditnota eller rykker
												parentfak = request("id")
												else
												parentfak = 0
												end if
													
													'***tidspunkt = 23:23:59 pga luk for indtastning muligheden på timeregsiden
													
													strSQL = ("INSERT INTO fakturaer"_
													&" (faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, tidspunkt, "_
													&" betalt, b_dato, fakadr, att, faktype, konto, modkonto, parentfak, "_
													&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail, "_
													&" visafsswift, visafsiban, visafscvr, moms, enhedsang, "_
													&" aftaleid, varenr, jobbesk, timersubtotal, visjoblog,"_
													&" visrabatkol, vismatlog, rabat, visjoblog_timepris, "_
													&" visjoblog_enheder, visafsfax, subtotaltilmoms, "_
													&" valuta, kurs, sprog, istdato, momskonto, visperiode, "_
													&" visjoblog_mnavn, jobfaktype, betbetint, brugfakdatolabel, "_
													&" istdato2, momssats, modtageradr, usealtadr, vorref, fak_ski, "_
													&" showmatasgrp, hidesumaktlinier, sideskiftlinier, labeldato, fak_abo, fak_ubv, visikkejobnavn) VALUES ("_
													&" '"& intFaknum &"',"_
													&" '"& fakDato &"',"_
													&" "& jobid &","_
													&" "& intTimer &","_
													&" "& intBeloeb &","_
													&" '"& strKomm &"',"_
													&" '"& strDato &"',"_
													&" '"& strEditor &"', '23:59:59', "_
													&" "& intFakbetalt &", '"& dtb_dato &"', "& intfakadr &", "_
													&" '"& strAtt &"', "& intType &", "& varKonto &", "_
													&" "& varModkonto &", "& parentfak &", "_
													&" "& intModland &", "& intModAtt &", "& intModTlf &", "_
													&" "& intModCvr &", "& intAfsTlf &", "& intAfsEmail &", "& intAfsSwift &", "_
													&" "& intAfsIban &", "& intAfsICVR &", "& intMoms &", "& intEnhedsang &", "_
													&" "& aftid &", '"& strVarenr &"', '"& jobBesk &"', "& timersubtotal &", "_
													&" "& visjoblog &", "& visrabatkol &", "& vismatlog &", "& intRabat &", "_
													&" "& visjoblog_timepris &", "& visjoblog_enheder &", "_
													&" "& intAfsFax &", "& subtotaltilmoms &", "& valuta &", "_
													&" "& kurs &", "& intSprog &", '"& istDato &"', "& momskonto &", "_
													&" "& visperiode &", "& visjoblog_mnavn &", "& jobfaktype &", "_
													&" "& betbetint &", "& brugfakdatolabel &", '"& istDato2 &"', "_
													&" "& momssats &", '"& modtageradr &"', "& usealtadr &", '"& vorref &"', "_
													&" "& fak_ski &", "& showmatasgrp &", "& hidesumaktlinier &", "& sideskiftlinier &", "_
													&" '"& labelDato &"', "& fak_abo &", "& fak_ubv &", "& visikkejobnavn &")")
													
													'Response.Write "subtotaltilmoms: " & subtotaltilmoms & " intMoms: "& intMoms &"<br>"
													'Response.write strSQL & "<br><br>"
													'Response.flush
													oConn.execute(strSQL)
													
													'Response.end
													
										
													
													''** Henter fak id ***
													strSQL3 = "SELECT Fid FROM fakturaer"
													oRec3.open strSQL3, oConn, 3
													oRec3.movelast
													if not oRec3.EOF then
													thisfakid = cint(oRec3("Fid")) 
													end if 
													oRec3.close	
													
													'*** Shadowcopy ***'
													'*** Hvis der oprettes en faktura på en aftale, hvor der er tilknyttet job, skal der oprettes en shadowcopy 
													'*** På de job der er tilknyttet så de kan lukkes for indtastning, samt at man kan se at der ligger en GHOSTET
													'*** Faktura på jobbet 
													
													if aftid <> 0 then
													job_tilknyttet_aftale = split(request("FM_job_tilknyttet_aftale"),",")
													for i = 0 TO UBOUND(job_tilknyttet_aftale, 1)
													    if i > 0 then
													    
													    strSQLshadowCopy = "INSERT INTO fakturaer (faknr, fakdato, jobid, shadowcopy, istdato) VALUES "_
													    &"('"& intFaknum &"', '"& fakDato &"', "& job_tilknyttet_aftale(i) &", 1, '"& istDato &"')"
													    
													    
													    'Response.Write strSQLshadowCopy
													    'Response.end
													    oConn.execute(strSQLshadowCopy)
													    end if
													next
													    
													
													end if
													
									
									
							end if	'** Opret
									
									
												
						
							if jobid <> 0 then '** Bruges kun hvis der oprettes faktura på job ***					
							
							
							'***********************************************************************************
							'*** Hvis faktura allerede er oprrettet en gang i denne session ***
							'***********************************************************************************
								
								 
								'************************************************************** 
								'*** Indsætter akt i fak_det 
								'**************************************************************
								if func = "dbred" then
								oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
								thisfakid = id
								end if
								
								'***** n starter på 1 **************
								'***** Aktivitets udspecificering ************
								for intcounter = 0 to antalAkt - 1
								
								thisAktId = request("aktId_n_"&intcounter&"")
									
									'** antal aktiviter udskrevet pga. forskellige timepriser 
									'antalsumaktprakt = request("antal_subtotal_akt_"&intcounter&"") 
									antalsumaktprakt = request("antal_n_"&intcounter&"") 
									
									for intcounter3 = -(antalsumaktprakt) to antalsumaktprakt
									
												
												'*** Enhedsprisen på denn akt.
												if len(request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")) <> 0 then
												enhpris = SQLBless2(request("FM_enhedspris_"& intcounter &"_"&intcounter3&""))
												else
												enhpris = 0
												end if
												
												'*** Enheds angivelse på denn akt.
												enhedsang_akt = request("FM_akt_enh_"& intcounter &"_"&intcounter3&"")
												
												
									
										'*** Vis sum aktivitet på print (og DB)
										if len(request("FM_show_akt_"& intcounter &"_"&intcounter3&"")) = 1 then
												
												timerThis = request("FM_timerthis_"& intcounter &"_"&intcounter3&"")
												if len(timerThis) <> 0 then
												timerThis = SQLBless2(timerThis)
												else 
												timerThis = 0
												end if
												
												if len(request("FM_beloebthis_"& intcounter &"_"&intcounter3&"")) <> 0 then
												beloebThis = SQLBless2(request("FM_beloebthis_"& intcounter &"_"&intcounter3&""))
												else
												beloebThis = 0
												end if
												
												'*** Rabat *****
												if len(request("FM_rabat_"&intcounter&"_"&intcounter3&"")) <> 0 then
												rabatThis = SQLBless2(request("FM_rabat_"& intcounter &"_"&intcounter3&""))
												else
												rabatThis = 0
												end if	
												
												'*** Valuta *****'
												if len(request("FM_valuta_"&intcounter&"_"&intcounter3&"")) <> 0 then
												valutaThis = SQLBless2(request("FM_valuta_"& intcounter &"_"&intcounter3&""))
												else
												valutaThis = 1
												end if	
												
												'**** Sortorder ***'
												aktsortorder = request("aktsort_"&intcounter&"")
												call erDetInt(trim(aktsortorder))
			                                    if isInt > 0 then
			                                    aktsortorder = 1
			                                    else
			                                    aktsortorder = replace(aktsortorder, ",", ".")
			                                    end if
												
												kursThis = replace(request("valutakurs_"& valutaThis &""), ",", ".")
												
												momsfri = 0
												if len(trim(request("FM_momsfri_"& intcounter &"_"&intcounter3&""))) <> 0 then
												momsfri = 1
												else
												momsfri = 0
												end if
												
												strSQL_sumakt = ("INSERT INTO faktura_det "_
												&" (antal, beskrivelse, aktpris, fakid, "_
												&" enhedspris, aktid, showonfak, rabat, enhedsang, valuta, kurs, fak_sortorder, momsfri) "_
												&" VALUES ("& timerThis &", "_
												&"'" & SQLBlessK(request("FM_aktbesk_"& intcounter &"_"&intcounter3&"")) &"', "_
												&""  & beloebThis &", "_
												&""  & thisfakid &", "& enhpris &", "& thisAktId &", "_
												&" 1, "& rabatThis &", "& enhedsang_akt &", "& valutaThis &", "& kursThis &", "& aktsortorder &", "& momsfri &")")
												
												'Response.write strSQL_sumakt  & "<br>"
												'Response.end
												oConn.execute(strSQL_sumakt)
												
										end if 'show sumaktivitet
													
										
										'**************************************************'			
										'***** Medarbejder udspecificering i db ***********'
										'**************************************************'
										
										
										'Response.write "5<br>"
										'Response.flush
											
										antalmedspec2 = request("antal_n_"&intcounter&"") 'medarbejdere
										for intcounter2 = 0 to antalmedspec2 - 1
												
												'*** Passer timeprisen på denne akt og medarbejder **
												thismedarbtpris = request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")
												thismedarbtpris = formatnumber(thismedarbtpris, 2)
												thismedarbtpris = replace(thismedarbtpris, ".", "")
												thismedarbtpris = SQLBless2(thismedarbtpris) 
												
												thisenhpris = request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")
												thisenhpris = formatnumber(thisenhpris, 2)
												thisenhpris = replace(thisenhpris, ".", "")
												enhpris = SQLBless2(thisenhpris)
												
												
												
															
												
												'** registrerer timer (fak og vent) på alle medarbejdere 
												if len(request("FM_mid_"&intcounter2&"_"&intcounter&"")) <> 0 then
												thisMid = request("FM_mid_"&intcounter2&"_"&intcounter&"")
												else
												thisMid = 0
												end if 
												
												'Response.Write "<hr>thisMid " & thisMid & "<br>" 
												
												'if thisMid <> 0 then
												'Response.write "" & request("FM_aktbesk_"& intcounter &"_"&intcounter3&"") & "<br>"
												'Response.write "if" & thismedarbtpris &" = "& enhpris  &" then<br>"
												'Response.write "timer:" & request("FM_m_fak_"&intcounter2&"_"&intcounter&"")
												'Response.write "<br>true sum aktr:" & request("FM_show_akt_"&intcounter&"_"&intcounter3&"")
												'end if
												
												if cdbl(thismedarbtpris) = cdbl(enhpris) AND thisMid <> 0 then
												'Response.write "<br>ok!<br>"
											
													
														'* Nulstiller evt. tidligere indtastninger på denne faktura ***
														oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& thisfakid &" AND aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														'* Beløb
														if len(request("FM_mbelob_"& intcounter2 &"_"&intcounter&"")) <> 0 then
														useBeloeb = SQLBless2(request("FM_mbelob_"& intcounter2 &"_"&intcounter&""))
														else
														useBeloeb = 0
														end if
														
														'* Venter
														if len(request("FM_m_vent_"&intcounter2&"_"&intcounter&"")) <> 0 then
														useVenter = SQLBless2(request("FM_m_vent_"&intcounter2&"_"&intcounter&""))
														else
														useVenter = 0
														end if
														
														'** Venter brugt 
														if len(request("FM_m_venterbrugt_"&intcounter2&"_"&intcounter&"")) <> 0 then
														useVenterBrugt = SQLBless2(request("FM_m_venterbrugt_"&intcounter2&"_"&intcounter&""))
														else
														useVenterBrugt = 0
														end if
														
														
														
														'*** Rabat *****'
														if len(request("FM_mrabat_"&intcounter2&"_"&intcounter&"")) <> 0 then
														medarb_rabat = SQLBless2(request("FM_mrabat_"& intcounter2 &"_"&intcounter&""))
														else
														medarb_rabat = 0
														end if	
														
														
														'*** Valuta *****'
												        if len(request("FM_mvaluta_"&intcounter2&"_"&intcounter&"")) <> 0 then
												        medarbValuta = SQLBless2(request("FM_mvaluta_"& intcounter2 &"_"&intcounter&""))
												        else
												        medarbValuta = 1
												        end if	
        												
												        medarbKurs = replace(request("valutakurs_"& medarbValuta &""), ",", ".")
														
														
														'*** Enhedsang ***
														enhedsang_med = request("FM_med_enh_"&intcounter2&"_"&intcounter&"")
														
														'** Nulstiler altid vente timer inden der tildeles nye vente timer for denne medarbejder på denne aktivitet 
														'** (Uanset hvilken fak akt. hører til)
														
														'oConn.execute("UPDATE fak_med_spec SET venter = 0 WHERE aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														
														if len(request("FM_m_fak_"&intcounter2&"_"&intcounter&"")) <> 0 then
															'*** Hvis show sum-akt ikke er true skal faktimer altid være = 0
															'if request("FM_show_akt_"&intcounter&"_"&intcounter3&"") = "1" then 
															useFak = SQLBless2(request("FM_m_fak_"&intcounter2&"_"&intcounter&""))
															'else
															'useFak = 0
															'end if
														else
															useFak = 0
														end if
														
														
														if len(request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")) <> 0 then
														usemedTpris = SQLBless2(request("FM_mtimepris_"& intcounter2&"_"&intcounter&""))
														else
														usemedTpris = 0
														end if 
														
														
														if request("FM_show_medspec_"&intcounter2&"_"&intcounter&"") = "show" then
														showonfak = 1
														else
														showonfak = 0
														end if
														
														
														
														'*** Indsætter i db ****
														if useFak <> 0 OR useVenter <> 0 then
														
														strSQL = "INSERT INTO fak_med_spec (fakid, aktid, mid, fak, venter, "_
														&" tekst, enhedspris, beloeb, showonfak, medrabat, venter_brugt, enhedsang, valuta, kurs) "_
														&" VALUES ("& thisfakid &", "&thisAktId&", "_
														&"" &thisMid&", "_
														&"" &useFak&", "_
														&"" &useVenter&", "_
														&"'"&SQLBlessK(request("FM_m_tekst_"& intcounter2&"_"&intcounter&""))&"', "_
														&"" &usemedTpris&", "_
														&"" &useBeloeb&", "& showonfak &", "_
														&"" & medarb_rabat &", "& useVenterBrugt &", "& enhedsang_med &", "& medarbValuta &", "& medarbKurs &")"
														
														
														'Response.write strSQL & "<br><br>"
														'Response.flush
														
														oConn.execute(strSQL)
													    
													    
													    end if
													    
													'else
													'Response.write "Stemmer ikke <br>"	
													
													end if 'thismedarbtpris
													
											next 'Medarbejdere
										next 'Sumaktiviteter
								next 'Antal aktiviteter
							
							
							'Response.write "6<br>"
							'Response.flush
							
							else	
							
							'select case intFakbetalt
							'case 1
							'erfak = 1 'godkendt
							'case 0
							'erfak = 2 'kladde
							'end select
							'*** Opdaterer seraft, så aftalen ER faktureret ***
							'oConn.execute("UPDATE serviceaft SET erfaktureret ="& erfak &" WHERE id = "& aftid &"")
									
							end if '*** Kun job (opr. aktiviterer og fak_med_spec)	
							
							
							
							'Response.end
							
							
							'********* Materialer ************'
							
							
							mat_ialt = request("FM_antal_materialer_ialt")
							
							
							
							if func = "dbred" then
							    '** renser db, opdater materiale forbrug ***'
							    strSQLselmf = "SELECT matfrb_id FROM fak_mat_spec WHERE matfakid = " & thisfakid
							    oRec4.open strSQLselmf, oConn, 3
							    while not oRec4.EOF
    							
							    '*** Markerer i materialeforbrug at materiale er faktureret ***'
							    strSQLmf = "UPDATE materiale_forbrug SET erfak = 0 WHERE id = " & oRec4("matfrb_id")
							    oConn.execute(strSQLmf)
    							
							    oRec4.movenext
							    wend
							    oRec4.close
    							
    							
							    '*** Sletter hidtidige regs ***'
							    strSQLdel = "DELETE FROM fak_mat_spec WHERE matfakid = " & thisfakid
							    oConn.execute strSQLdel
							end if
							
							
							
							for m = 0 to mat_ialt - 1
							
							
							if len(trim(request("FM_vis_"&m&""))) <> 0 then
							matvis = 1
							else
							matvis = 0
							end if
							
							matid = request("FM_matid_"&m&"")
							matnavn = replace(request("FM_matnavn_"&m&""),"'", "''")
							matvarenr = request("FM_matvarenr_"&m&"")
							matantal = SQLBless2(request("FM_matantal_"&m&""))
							matenhed = request("FM_matenhed_"&m&"")
							matenhedspris = SQLBless2(request("FM_matenhedspris_"&m&""))
							matrabat = request("FM_matrabat_"&m&"")
							matbeloeb = SQLBless2(request("FM_matbeloeb_"&m&""))
							matValuta = request("FM_matvaluta_"&m&"")
							matKurs = replace(request("valutakurs_"& matValuta &""), ",", ".")
							matGrp = request("FM_matgrp_"&m&"")
							
							if len(trim( request("FM_matikkemoms_"&m&""))) <> 0 then
							ikkemoms = request("FM_matikkemoms_"&m&"")
							else
							ikkemoms = 0
							end if
							
							matMFid = request("FM_mfid_"&m&"")
							matMFusrid = request("FM_mfusrid_"&m&"")
							
							if cint(matvis) = 1 then
							
							strSQLoprmat = "INSERT INTO fak_mat_spec (matfakid, matid, matnavn, matvarenr, matantal, matenhed, "_
							&" matenhedspris, matrabat, matbeloeb, matshowonfak, ikkemoms, valuta, kurs, matfrb_mid, matfrb_id, matgrp) "_
							&" VALUES ("&thisfakid&", "&matid&", '"&matnavn&"', '"& matvarenr&"', "_
							&" "& matantal &", '"& matenhed &"', "& matenhedspris &", "_
							&" "& matrabat &", "& matbeloeb &", 1, "& ikkemoms &", "& matValuta &", "& matKurs &", "& matMFusrid &" ,"& matMFid &", "& matGrp &")"
							
							'Response.Write strSQLoprmat & "<br>"
							'Response.flush
							oConn.execute(strSQLoprmat)
							
							'matbeloebTot = matbeloebTot + matbeloeb
							
							'*** Markerer i materialeforbrug at materiale er faktureret ***'
							strSQLmf = "UPDATE materiale_forbrug SET erfak = 1 WHERE id = " & matMFid
							oConn.execute(strSQLmf)
							
							end if
							    
							next
							
							
							
							
			                '**********************************'
			                '************************************'
							
							
					         '**** Opdaterer fakturanr i licens tabel ****'
					        if func = "dbopr" then
                            strSQL = "UPDATE licens SET "& sqlfld &" = "& intFaknum &" WHERE id = 1"
                            oConn.execute(strSQL)
                            end if
    					
							'*** Lukker job ****'
							if len(request("FM_lukjob")) <> 0 then
							strSQLupd = "UPDATE job SET jobstatus = 0 WHERE id = " & jobid
							oConn.execute(strSQLupd)
							end if
						
							
							'**** Opdaterer faste betalingsbetingelser på kunde ***'
							if len(request("FM_gembetbet")) <> 0 then
							strSQLupd = "UPDATE kunder SET betbet = '"& strKomm &"', betbetint = "& betbetint &" WHERE kid = " & kundeid
							oConn.execute(strSQLupd)
							end if
							
							
							'** Gemmer på alle kunder **'
							if len(request("FM_gembetbetalle")) <> 0 then
							strSQLupd = "UPDATE kunder SET betbet = '"& strKomm &"', betbetint = "& betbetint &" WHERE kid <> 0"
							oConn.execute(strSQLupd)
							end if
							
							
							
							
							
							'*** Opretter posteirnger på Execon / Immenso version ***'
							if intFakbetalt <> 0 then 
							    call ltoPosteringer
							end if
							
							'*** Opdaterer/Indsætter posterings ID på faktura ***' 
							'*** Må først kaldes efer alle posteringer ***'
							if intFakbetalt <> 0 then 
							strSQLOprID = "UPDATE fakturaer SET oprid = "& oprid &" WHERE Fid = "& thisfakid &""
							oConn.execute(strSQLOprID)
	                        end if
							
							
							Response.Cookies("erp").Expires = date + 60
							
							'Response.End 
							
												
												
											'*** Viser den oprettede faktura til print *****'
											%>
											<!-------------------------------Sideindhold------------------------------------->
											
										
											<%
											stDag = request("FM_start_dag_ival")
											stMrd = request("FM_start_mrd_ival")
											stAar = request("FM_start_aar_ival")
											
											slDag = request("FM_slut_dag_ival")
											slMrd = request("FM_slut_mrd_ival")
											slAar = request("FM_slut_aar_ival")
											
											
											
											
											%>
											 
											
											<script>
											
									        window.top.frames['erp3'].location.href = "erp_fak_godkendt_2007.asp?jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=thisfakid%>&FM_usedatointerval=1&FM_start_dag_ival=<%=stDag%>&FM_start_mrd_ival=<%=stMrd%>&FM_start_aar_ival=<%=stAar%>&FM_slut_dag_ival=<%=slDag%>&FM_slut_mrd_ival=<%=slMrd%>&FM_slut_aar_ival=<%=slAar%>&kid=<%=kid%>"
	                                        window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?FM_type=<%=intType%>&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=stDag%>&FM_start_mrd=<%=stMrd%>&FM_start_aar=<%=stAar%>&FM_slut_dag=<%=slDag%>&FM_slut_mrd=<%=slMrd%>&FM_slut_aar=<%=slAar%>"
	                                        
	                                        </script>
											
											<%
											
											
									'end if faknrfindesallerede
							end if
							
						end if
					end if
			end if
	
	
	
	
	
	
	
	
	
	
	
	case else
	'**************************** Opret/rediger Faktura Data *********************************'
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<SCRIPT language=javascript src="inc/fak_func_2007.js"></script>
	<!--#include file="inc/fak_inc_subs_2007.asp" -->
	<!-------------------------------Sideindhold------------------------------------->
	<%
	
	
	



	
	
	'******************************************************
	'*** Opret / Rediger faktura **************************
	'******************************************************
	
	
	
	timeforbrug = 0
	faktureretBeloeb = 0
	
	pleft = 0
	ptop = 0
	
	Redim thisaktid(115)
	Redim thisaktnavn(115)
	Redim thisAktTimer(115)
	Redim thisAktBeloeb(115)
	'Redim usefastpris(115)
	Redim thisTimePris(115)
	Redim thisaktfunc(115)
	Redim thisaktfakbar(115)
	Redim thisaktfaktor(115)
	Redim thisaktsort(115)
	
	Redim thisaktrabat(115)
	
	Redim thisaktbesk(115)
	Redim thisaktForkalk(115)
	redim thisAktForkalkStk(115)
	redim thisAktBudgetsum(115)
	
	Redim thisAktEnhAng(115)
	Redim thisAktEnhPris(115)
	
	Redim thisAktValuta(115)
	Redim thisAktText(115)
	
	Redim thisaktCHK(115) 
	Redim thisaktEnh(115)
	
	redim thisAktFase(115)
	redim thisMomsfri(115)
	
	redim thisAktBgr(115)
	redim thisJoborAkt_tp(115)
	
	if func = "red" then
	'*********************************************************************
	' Rediger faktura 
	'*********************************************************************
	
	
						strSQL = "SELECT Fid, faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, betalt, b_dato, fakadr, "_
						&" att, faktype, konto, modkonto, varenr, enhedsang, jobbesk, "_
						&" visjoblog, visrabatkol, vismatlog, visjoblog_timepris, "_
						&" visjoblog_enheder, visjoblog_mnavn, rabat, visafsemail, visafsswift, visafstlf, "_
						&" visafsiban, visafscvr, vismodland, vismodatt, vismodtlf, "_
						&" vismodcvr, visafsfax, valuta, kurs, sprog, istdato, momskonto, visperiode, jobfaktype, "_
						&" betbetint, brugfakdatolabel, istdato2, "_
						&" momssats, modtageradr, usealtadr, vorref, fak_ski, showmatasgrp, hidesumaktlinier, "_
						&" sideskiftlinier, labeldato, fak_abo, fak_ubv, visikkejobnavn"_
						&" FROM fakturaer WHERE Fid = "& id 
						
						'Response.Write strSQL
						'Response.Flush
						
						oRec.open strSQL, oConn, 3
						
						if not oRec.EOF then
						strEditor = oRec("editor")
						strFaknr = oRec("faknr")
						strTdato = formatdatetime(oRec("fakdato"), 2)
						'fakdatovente = oRec("fakdato")
						
						
						visafsemail = oRec("visafsemail")
						visAfsSwift = oRec("visafsswift")
					    visAfsIban = oRec("visafsiban")
					    visAfsCVR = oRec("visafscvr")
						visAfsTlf = oRec("visafstlf")
						visAfsFax = oRec("visafsfax")
						
						    intModland = oRec("vismodland")
						
						    if cint(intModland) <> 0 then
			                'strLand =  "<br>"& oRec("land")
			                'strLandShow = strLand
			                lndCHK = "CHECKED"
			                else
			                'strLandShow = "Danmark"
			                'strLand = ""
			                lndCHK = ""
			                end if
						
						
						
						'intModTlf = oRec("vismodtlf")
						intModCvr = oRec("vismodcvr")
							
								
						dbfunc = "dbred"		
						strKom = oRec("kommentar")
								
						strTimer = oRec("timer")
						strBeloeb = oRec("beloeb")
						'StrUdato = "12/12/2014"
						strDato = oRec("dato")
						intBetalt = oRec("betalt")
						strB_dato = oRec("b_dato")
						
						
						intFakid = oRec("Fid")
						intfakadr = oRec("fakadr")
						strAtt = oRec("att")
						
						intType = oRec("faktype")
						
						intKonto = oRec("konto")
						intModKonto = oRec("modkonto")
						
						strVarenr = oRec("varenr")
						intEnhedsang = oRec("enhedsang")
						
						strJobBesk = oRec("jobbesk")
						strKom = oRec("kommentar")
						
						visjoblog = oRec("visjoblog")
						visrabatkol = oRec("visrabatkol")
						vismatlog = oRec("vismatlog")
						
						visjoblog_timepris = oRec("visjoblog_timepris")
						visjoblog_enheder = oRec("visjoblog_enheder") 
						visjoblog_mnavn = oRec("visjoblog_mnavn")
						
						intRabat = oRec("rabat")
						
						valuta = oRec("valuta")
						kurs = oRec("kurs")
						valutaKursFak = kurs
						sprog = oRec("sprog")
						
						istDato = oRec("istdato")
						istDato2 = oRec("istDato2")
						
						momskonto = oRec("momskonto")
						visperiode = oRec("visperiode")
						
						jobfaktype = oRec("jobfaktype")
						
						betbetint = oRec("betbetint")
						brugfakdatolabel = oRec("brugfakdatolabel")
						
						strForfaldsdato = oRec("b_dato")
						
						momssats = oRec("momssats")
						modtageradr = oRec("modtageradr")
						
						usealtadr = oRec("usealtadr")
						
						vorref = oRec("vorref")
						
						ski = oRec("fak_ski")
						abo = oRec("fak_abo")
						ubv = oRec("fak_ubv")
						
						showmatasgrp = oRec("showmatasgrp")
						
						hidesumaktlinier = oRec("hidesumaktlinier")
						sideskiftlinier = oRec("sideskiftlinier")
						
						labelDato = oRec("labeldato")
						
						visikkejobnavn = oRec("visikkejobnavn")
						
						end if
						oRec.close
						
						
						if len(intModKonto) <> 0 then
						intModKonto = intModKonto
						else
						intModKonto = 0
						end if
						
						if intBetalt = 1 then
						betaltch = "checked"
						else
						betaltch = ""
						end if 
						
						
	    
	    
	                if jobid <> 0 then
            							
							            call faktureredetimerogbelob()
            							
            						
							            strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn, "_
							            &" ikkebudgettimer, jobans1, jobans2, rekvnr, jobstatus, usejoborakt_tp, fastpris FROM job WHERE id = " & jobid
							            oRec.open strSQL, oConn, 3
							            if not oRec.EOF then
            								
            								
            								
								            intIkkeBtimer = oRec("ikkebudgettimer")
								            intBudgettimer = oRec("budgettimer")
								            intJobTpris = oRec("jobTpris")
								            fastpris = oRec("fastpris")
								            if intfakadr <> 0 then
								            kid = intfakadr
								            else
								            kid = oRec("jobknr")
								            end if
								            intjobnr = oRec("jobnr")
								            strJobnavn = oRec("jobnavn")
            								
								            jobans1 = oRec("jobans1")
								            jobans2 = oRec("jobans2")
								           
            								rekvnr = oRec("rekvnr")
            								jobstatus = oRec("jobstatus")
            								
            								usejoborakt_tp = oRec("usejoborakt_tp")
            								
            								usefastpris = oRec("fastpris")
            								
            								
            								
							            end if
							            oRec.close
            							
		            else
            		
			            kid = intfakadr
            			
			            intEnheder = strTimer
			            intPris = strBeloeb
			            strBesk = strKom
            		
		            end if
		'*******************************************************************************************
	    
	 oimg = "ac0010-24.gif"
	oleft = 0
	otop = -20
	owdt = 400
	
	if cint(intType) <> 1 then
	typTxt = "Faktura"
	else
	typTxt = "Kreditnota"
	end if
	
	oskrift = "Rediger "& typTxt &" "& strFaknr 
		
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
		
		'** Mulighed for at slette og se sidst redigeret dato ****'
		%>
		<div id="sidetop" style="position:absolute; left:0px; top:5px; visibility:visible; z-index:100;">
	    
	    
	     <div id=sidstopd style="position:absolute; left:5px; width:350px; top:22px; border:0px #8caae6 solid;">
	        <table cellspacing="0" cellpadding="0" border="0" width="370"><tr>
				<td height="25">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b>
				<!--&nbsp;&nbsp;<a href="fak.asp?menu=stat_fak&func=slet&id=<=id%>&FM_job=<=Request("FM_job")%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="Slet denne faktura" border="0"></a>-->
				</td>
			</tr>
			</table>
			</div>
		
		
		
		<% 
	
	
	
	
	
	else
	'************************************************************************'
	' Opret faktura 
	'************************************************************************'
	
	oimg = "ac0010-24.gif"
	oleft = 0
	otop = -20
	owdt = 400
	
	if cint(intType) <> 1 then
	typTxt = "Faktura"
	else
	typTxt = "Kreditnota"
	end if
	oskrift = "Opret ny " & typTxt
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	%>
	
	<!-- main -->
	
	<div id="sidetop" style="position:absolute; left:0px; top:5px; visibility:visible; z-index:100;">
    
    
    
   
	
	
	<%
	
	
		
		
		
		
		'*** Find kunde vedr opret faktura på JOB ***
		if jobid <> 0 then
		
		call faktureredetimerogbelob()
		
		strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, "_
		&" jobnavn, ikkebudgettimer, jobans1, jobans2, kundekpers, beskrivelse, valuta, "_
		&" jobstartdato, jobslutdato, jobfaktype, rekvnr, jobstatus, usejoborakt_tp, ski, abo, ubv FROM job WHERE id = " & jobid
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
			
			intIkkeBtimer = oRec("ikkebudgettimer")
			intBudgettimer = oRec("budgettimer")
			intJobTpris = oRec("jobTpris")
			fastpris = oRec("fastpris")
			
			kid = oRec("jobknr")
			intjobnr = oRec("jobnr")
			strJobnavn = oRec("jobnavn")
			
			jobans1 = oRec("jobans1")
			jobans2 = oRec("jobans2")
			
			intKundekpers = oRec("kundekpers")
			strJobBesk = oRec("beskrivelse")
			
			valuta = oRec("valuta")
			jobstdato = oRec("jobstartdato")
			jobsldato = oRec("jobslutdato")
			
			'if len(request("FM_jobfaktype")) <> 0 then
			'jobfaktype = request("FM_jobfaktype")
			'Response.Write "<br><br>her" & jobfaktype
			'else
			'jobfaktype = oRec("jobfaktype")
			'end if
			
			jobfaktype = 0
			
			rekvnr = oRec("rekvnr")
			jobstatus = oRec("jobstatus")
			
			
			usejoborakt_tp = oRec("usejoborakt_tp")
			ski = oRec("ski")
			
			abo = oRec("abo")
			ubv = oRec("ubv")

			
			'jobtimerforkalk = intIkkeBtimer + intBudgettimer
			
		end if
		oRec.close
		
		else 
		
		'*** Find kunde vedr opret faktura på Aftale ***'
		
		call faktureredetimerogbelob()
		
		strSQL = "SELECT kundeid, besk, enheder, pris, varenr, navn, perafg, "_
		&" advitype, advihvor, stdato, sldato, valuta FROM serviceaft WHERE id = "& aftid
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		kid = oRec("kundeid")
		'strBesk = oRec("besk")
		intEnheder = oRec("enheder")
		intPris = oRec("pris")
		strVarenr = oRec("varenr")
		strNavn = oRec("navn")
		strJobBesk = strNavn
		intPerafg = oRec("perafg")
		intAdvitype = oRec("advitype")
		intAdvihvor = oRec("advihvor")
		startdato = oRec("stdato") 
		slutdato = oRec("sldato") 
		valuta = oRec("valuta")
		
		oRec.movenext
		wend
		oRec.close 
		
		
		jobfaktype = 0
		strBesk = strNavn&":&nbsp;"&strBesk
		
		
		end if '*** Aft / Job **'
		
		
		kurs = ""
		sprog = 1
		
		strAtt = intKundekpers
		
		intKonto = 0
		intModKonto = 0
		intType = intType 
		visjoblog = chklog
		
		
		
		
		'*** Skal rabat kolonne være slået til default. ***'
		select case lto
		case ""
		visrabatkol = 1
		case else
		visrabatkol = 0
		end select 
		
		
		vismatlog = chklog
		'lastFaknr = 0
		
		visjoblog_timepris = 1
		visjoblog_enheder = 1
		visjoblog_mnavn = 1
		brugfakdatolabel = 0
		
		
		strEditor = ""
		strTdato = "" 'month(now)&"/"&day(now)&"/"& year(now)
		istDato = ""
		istDato2 = ""
		
		strKom = ""
		dbfunc = "dbopr"
		varSubVal = "Opretpil" 
		
		if len(chkemail) <> 0 then
		chkemail = chkemail
		else
		chkemail = 0 
		end if
		
		if cint(chkemail) = 1 then
		visafsemail = 1
		else
		visafsemail = 0
		end if
		
		visAfsSwift = 0
		visAfsIban = 0
	    visAfsCVR = 1
		
		if len(chktlffax) <> 0 then
		chktlffax = chktlffax
		else
		chktlffax = 0
		end if
		
		if cint(chktlffax) = 1 then
		visAfsTlf = 1
		visAfsFax = 1
		else
		visAfsTlf = 0
		visAfsFax = 0
		end if
		
		'intModland = 0
		'intModTlf = 0
		intModCvr = 0
		
		momskonto = 1
		momssats = 25
		visperiode = 0
		
		usealtadr = 0
		vorref = "-"
		showmatasgrp = 0
		
		hidesumaktlinier = 0
		
		select case lto
	    case "dencker", "tooltest", "outz", "intranet - local", "optionone"
	    sideskiftlinier = 11
	    case else
	    sideskiftlinier = 0
	    end select
		
		
		visikkejobnavn = 0
				
		'*****************************************'
		'*****************************************'
	end if
	
	
	if visafsemail = 1 then
	visafsemailCHK = "CHECKED"
	else
	visafsemailCHK = ""
	end if
	
	if visAfsTlf = 1 then
	visAfsTlfCHK = "CHECKED"
	else
	visAfsTlfCHK = ""
	end if
	
	if visAfsFax = 1 then
	visAfsFaxCHK = "CHECKED"
	else
	visAfsFaxCHK = ""
	end if
	
	if visAfsIban = 1 then
	visAfsIbanCHK = "CHECKED"
	else
	visAfsIbanCHK = ""
	end if
	
	if visAfsSwift = 1 then
	visAfsSwiftCHK = "CHECKED"
	else
	visAfsSwiftCHK = ""
	end if
	
	if visafsCVR = 1 then
	visafscvrCHK = "CHECKED"
	else
	visafscvrCHK = ""
	end if
	
	'if intModland = 1 then
	'intModlandCHK = "CHECKED"
	'else
	'intModlandCHK = ""
	'end if
	
	'if intModTlf = 1 then
	'intModTlfCHK = "CHECKED"
	'else
	'intModTlfCHK = ""
	'end if
	
	if intModCvr = 1 then
	intModCvrCHK = "CHECKED"
	else
	intModCvrCHK = ""
	end if
	
	
	if visikkejobnavn <> 0 then
	visikkejobnavnCHK = "CHECKED"
	else
	visikkejobnavnCHK = ""
	end if
	
	%>
	
	                
	                
	                <!-- Valuta load -->

                      <div id="valutaload" style="position:absolute; top:61px; left:125px; height:200px; display:none; visibility:hidden; border:2px red dashed; background-color:#FFFFe1; padding:20px; width:300px; height:150px; z-index:20000000;">
	                                  <b>Omregner priser til valgte valuta, et øjeblik...</b><br /><br />
                                          <img src="../ill/loaderbar.gif" /><br /><br />
                                         Det tager normalt 5-8 sek.<br />
                          &nbsp;
                     </div>
	               
	               
	               <!-- Sideload -->
	              <div id="sideload" style="position:absolute; top:61px; left:125px; height:200px; display:; visibility:visible; border:2px red dashed; background-color:#FFFFe1; padding:20px; width:300px; height:150px; z-index:20000;">
	              <b>Beregner side, et øjeblik...</b><br /><br />
                      <img src="../ill/loaderbar.gif" /><br /><br />
                      Hvis der er valgt et langt dato interval, eller hvis der er mange medarbejder- og aktivitets -linier, kan det lidt tid at vise siden.
	                    <br /><br />
	                    
	                    
	                    
	                    <%
	                    antalaktCount = 0
	                    strSQL = "SELECT COUNT(id) AS antalaktCount FROM aktiviteter WHERE job = "& jobid &" GROUP BY job" 
	                    
	                    'Response.Write strSQL
	                    'Response.flush
	                    
	                    oRec.open strSQL, oConn, 3
	                    if not oRec.EOF then
	                    antalaktCount = oRec("antalaktCount")
	                    end if
	                    oRec.close 
	                    %>
	                    
	                    Der er <b><%=antalaktCount %></b> aktiviter på dette job. Hver aktivitet tager 0,2-2 sek. at loade.
	                    Dvs. at det ca. tager mellem:<br /> <b><%=formatnumber((antalaktCount*0.2), 0) %> og <%=formatnumber((antalaktCount*1.4), 0) %> </b> sekunder at vise siden.
	                    
	                   
	                    <%'Response.flush %>
	              </div>
	                
	                
	                
	                
	                
	                <!-- Menu -->
	                
	                <div id=menu style="position:absolute; top:31px; height:65px; left:5px; display:none; visibility:hidden;">
	                
	                <div id=knap_joblogdiv style="position:absolute; visibility:visible; display:; top:-30px; width:100px; left:433px; border:1px #8cAAe6 solid; padding:3px 3px 3px 3px; background-color:#EFf3FF;">
                    <a href="#" onclick="showdiv('joblogdiv')" class=vmenu>Joblog</a>
	                </div>
	                
	                <div id=knap_matlogdiv style="position:absolute; visibility:visible; display:; top:-30px; width:150px; left:538px; border:1px #8cAAe6 solid; padding:3px 3px 3px 3px; background-color:#EFf3FF;">
                    <a href="#" onclick="showdiv('matlogdiv')"  class=vmenu>Materiale- / udlægs -log</a>
	                </div>
	                
	                
	                
	               
                    
                    <div id=knap_fidiv style="position:absolute; visibility:visible; display:; top:15px; width:95px; left:0px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffff99;">
                        <table cellspacing=0 cellpadding=0 border=0 width=100%>
                        <tr><td colspan=2 class=lille>Step 1</td></tr>
                        <tr>
                        <td><a href="#" onclick="showdiv('fidiv')" class=vmenu>Faktura<br /> indstillinger</a></td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	                
	                <%if jobid <> 0 then
	                wdt = "100"
	                else
	                wdt = "200"
	                end if%>
	                
	                <div id=knap_modtagdiv style="position:absolute; visibility:visible; display:; top:15px; width:<%=wdt%>px; left:95px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille>Step 2</td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('modtagdiv')" class=vmenu>Modtager og<br /> Afsender</a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	               
	                
	                <%
	                if jobid <> 0 then
	                %>
	                
	                <div id=knap_jobbesk style="position:absolute; visibility:visible; display:; top:15px; width:65px; left:195px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille>Step 3</td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('jobbesk')" class=vmenu>Job<br />beskriv.</a></td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	               
	                
	                
	             
	                
	                <div id=knap_aktdiv style="position:absolute; visibility:visible; display:; top:15px; width:90px; left:265px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille>Step 4</td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('aktdiv')" id="menushowakt" class=vmenu>Aktiviteter<br />fak. linier</a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	               
	                
	                
	                <div id=knap_matdiv style="position:absolute; visibility:visible; display:; top:15px; width:85px; left:355px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid;  padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille>Step 5</td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('matdiv')" class=vmenu>Materialer<br />fak. linier</a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	                
	                <%else%>
	                
	                <div id=knap_aktdiv style="position:absolute; visibility:visible; display:; top:15px; width:145px; left:295px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille>Step 3</td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('aktdiv')" class=vmenu>Aftale udspecifi. <br />Faktura linjer</a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	                
	                <%end if%>
	              
	                
	                
	                <div id=knap_betdiv style="position:absolute; visibility:visible; display:; top:15px; width:122px; left:440px; border:1px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <%
	                if jobid <> 0 then
	                 %>
	                 <tr><td colspan=2 class=lille>Step 6</td></tr>
                     <%else %>
                     <tr><td colspan=2 class=lille>Step 4</td></tr>
                     <%end if %>
                     <tr>
                        <td><a href="#" onclick="showdiv('betdiv')" class="vmenu">Betalings<br />betingelser</a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pile.gif" border=0 /> </td></tr></table>
	                </div>
	                
	                
	               
	                
	              </div>
	
	
	    <!-- gen indlæs faktura (interval ændring) KUN ved opret, ellers bruges intervl gemt med fak. -->
	             
	    <div id="genindlaes" style="position:absolute; top:192px; left:435px; z-index:2;">
		<form action="erp_opr_faktura_fs.asp?FM_kunde=<%=kid %>&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&reset=1&func=<%=func%>&FM_usedatokri=1&fid=<%=id%>" method="post" target="_top">
		
		
		
		<input type="hidden" name="FM_start_dag" id="gen_FM_start_dag" value="<%=request("FM_start_dag")%>">
		<input type="hidden" name="FM_start_mrd" id="gen_FM_start_mrd" value="<%=request("FM_start_mrd")%>">
		<input type="hidden" name="FM_start_aar" id="gen_FM_start_aar" value="<%=request("FM_start_aar")%>">
		<input type="hidden" name="FM_slut_dag" id="gen_FM_slut_dag" value="<%=request("FM_slut_dag")%>">
		<input type="hidden" name="FM_slut_mrd" id="gen_FM_slut_mrd" value="<%=request("FM_slut_mrd")%>">
		<input type="hidden" name="FM_slut_aar" id="gen_FM_slut_aar" value="<%=request("FM_slut_aar")%>">
		  
		  <%if func <> "red" then %> 
		  <input id="Button1" onclick="opdaterFakdato()" type="submit" value="Gen-indlæs faktura (hent timer i valgt. per.) >>" style="font-size:9px; width:225px;" /> 
	      <%end if %>
		
		</form>
		</div>
        
	
	
	
    
    
    
    <!--include file="inc/dato2.asp"-->
	<form name=main id=main action="erp_fak.asp?menu=stat_fak&func=<%=dbfunc%>" method="post" target="erp3">
	<%
	'*** Bruger altid datointerval ****
	usedt_ival = 1
	%>
	
	
	<input type="hidden" name="FM_usedatointerval" value="<%=usedt_ival%>">
	<input type="hidden" name="FM_start_dag_ival" value="<%=request("FM_start_dag")%>">
	<input type="hidden" name="FM_start_mrd_ival" value="<%=request("FM_start_mrd")%>">
	<input type="hidden" name="FM_start_aar_ival" value="<%=request("FM_start_aar")%>">
	<input type="hidden" name="FM_slut_dag_ival" value="<%=request("FM_slut_dag")%>">
	<input type="hidden" name="FM_slut_mrd_ival" value="<%=request("FM_slut_mrd")%>">
	<input type="hidden" name="FM_slut_aar_ival" value="<%=request("FM_slut_aar")%>">
	
	<input type="hidden" name="FM_fakint_ival" value="<%=trim(request("FM_fakint"))%>">
			
	
	<!--<input type="hidden" name="FM_type" value="<%=intType%>">-->
	<input type="hidden" name="FM_jobnavn" value="<%=strJobnavn%>">
	<input type="hidden" name="FM_job" id="FM_job" value="<%=Request("FM_job")%>">
	<input type="hidden" name="id" value="<%=id%>">
	
	<!--<input type="hidden" name="jobid" value="<%=jobid%>">-->
	
	<input type="hidden" name="FM_aftale" value="<%=aftid%>">
	<input type="hidden" name="FM_job_tilknyttet_aftale" value="<%=request("FM_job_tilknyttet_aftale")%>">
	<!--<input type="hidden" name="FM_aftnr" value="<%=sogaftnr%>">-->
	<input type="hidden" name="FM_kunde" value="<%=kid%>">
	
	<!--<input id="jobelaft" name="jobelaft" type="hidden" value="<%=jobelaft%>"/>-->
	<input id="faknr" name="faknr" type="hidden" value="<%=strFaknr%>"/>
	<input id="lastopendiv" name="lastopendiv" type="hidden" value="fidiv"/>
	
	<input id="jobfaktype" name="jobfaktype" type="hidden" value="<%=jobfaktype%>"/>
	
    
    
    <%call godkendknap() %>
	
	            
	 <%'** Valuta kurser **'
    strSQLval = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
    oRec.open strSQLval, oConn, 3
    while not oRec.EOF 
    %>
    <input id="valutakurs_<%=oRec("id")%>" name="valutakurs_<%=oRec("id")%>" value="<%=oRec("kurs") %>" type="hidden" />
    <%
    oRec.movenext
    wend
    oRec.close%>
	
	    
	    
	
	                <div id=modtagdiv style="position:absolute; visibility:hidden; display:none; top:105px; width:720px; left:5px; border:1px #8cAAe6 solid;">
                     <table cellspacing="0" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
                     <tr><td valign=top wstyle="width:350px;">
                     
                    <table cellspacing="0" cellpadding="5" border="0" width="100%">
	                <tr>
	                    <td style="border:1px #ffffff solid;" bgcolor="#8caae6" class=alt><b>Modtager </b></td>
	                </tr>
	                
                		<%
                		if func <> "red" then
		                strSQL = "SELECT kid, kkundenavn, kkundenr, adresse, "_
		                &" postnr, city, land, telefon, cvr, betbet, ktype, kt.rabat, ean, betbetint"_
		                &" FROM kunder LEFT JOIN kundetyper kt ON (kt.id = ktype) WHERE kid =" & kid  		
                		
		                'Response.Write strSQL
		                'Response.flush
                		
                		
		                oRec.open strSQL, oConn, 3
		                if not oRec.EOF then 
			                'intKnr = oRec("kkundenr")
			                strKnavn = trim(oRec("kkundenavn"))
			                strKadr = trim(oRec("adresse"))
			                strKpostnr = trim(oRec("postnr"))
			                strBy = trim(oRec("city"))
			                strLand = trim(oRec("land"))
			                intKid = oRec("Kid")
			                intCVR = oRec("cvr")
			                intTlf = oRec("telefon")
			                
			                if len(trim(oRec("ean"))) <> 0 then
			                ean = trim("<br>EAN: "&oRec("ean"))
			                else
			                ean = ""
			                end if
			             
                			
                			
                			betbetint = oRec("betbetint")
			                intRabat = oRec("rabat")
			                
			        
			                '*** Betalings betingelser ****
			                strKom = oRec("betbet")
			                
			                
			                if lcase(strLand) <> "danmark" then
			                strLand =  "<br>"& oRec("land")
			                strLandShow = strLand
			                lndCHK = "CHECKED"
			                else
			                strLandShow = "Danmark"
			                strLand = ""
			                lndCHK = ""
			                end if
                			
                			modtageradr = oRec("kkundenavn") & "<br>"& oRec("adresse") & "<br>"& oRec("postnr") & " " &oRec("city") & strLand & ean
           
                			
		                end if
                		
		                oRec.close
		                
		                else
		                modtageradr = modtageradr
		                betbetint = betbetint
			            intRabat = 100 * intRabat 
		                end if
		                
		                thisakt_rabat = intRabat
		                
		                %>
		                <tr>
		                <td valign=top style="height:110px;" align=left>
		                
                        <div id="DIV_modtageradr" style="width:310px; height:100px; padding:4px; border:1px #8CAAE6 solid; overflow:auto;">
                       <%=modtageradr%>
                        </div>
                            
                            <textarea style="position:absolute; visibility:hidden; width:310px; height:101px; padding:4px; border:1px #CCCCCC solid; top:30px; left:5px;" id="FM_modtageradr" name="FM_modtageradr">
                            <%=modtageradr%>
                            </textarea>   
                            <input type="button" id="gem_adr" style="position:absolute; display:none; visibility:hidden; top:30px; left:320px;" value="luk" />
                            
                            
                            <img src="../ill/blyant.gif" id="red_adr" border="0" style="position:absolute; display:; cursor:hand; visibility:visible; top:30px; left:320px;"/>
                            
		                 <input type="hidden" name="FM_Kid" id="FM_kid" value="<%=kid%>">
                		
		               </td>
		                </tr>
		               <tr>
		               <td><input type="checkbox" name="FM_land_vis" id="FM_land_vis" value="on" <%=lndCHK %>>Vis land: <%=strLandShow%>
		               <br />
		               <input type="checkbox" name="FM_cvr_vis" id="FM_cvr_vis" value="on" <%=intModCvrCHK %>>Vis CVR nr.: <%=intCVR%></td>	
		                </tr>
		                  <tr>	
		                <%if func <> "red" OR (func = "red" AND cint(strAtt) <> 0) then
		                attCHK = "CHECKED"
		                else
		                attCHK = ""
		                end if
		                 %>
			                <td>
				              
				                
				                <%if usealtadr <> 0 then
				                usealtadrCHK = "CHECKED"
				                else
				                usealtadrCHK = ""
				                end if %>
                                <input id="FM_usealtadr" name="FM_usealtadr" type="checkbox" value="1" <%=usealtadrCHK %> /> Benyt alternativ modtager adr. (Kontaktpers. / filial)<br />
                                <b>Kan ikke benyttes til offentlig fakturering via XML.</b> 
                                <br /><select name="FM_altadr" id="FM_altadr" style="width:150px; font-size:9px;">
				                <%
				                strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& kid &" ORDER BY navn"
				                oRec2.open strSQLkpersCount, oConn, 3
                					
					                while not oRec2.EOF
						                if cint(usealtadr) = cint(oRec2("id")) then
						                attSel = "SELECTED"
						                else
						                attSel = ""
						                end if%>
						                <option value="<%=oRec2("id")%>" <%=attSel%>><%=oRec2("navn")%></option>
						                <%
					                oRec2.movenext
					                wend
					                oRec2.close
					                %>
				                
				                </select>&nbsp;<input id="nyekpers" type="button" style="font-size:9px; font-family:arial;" value="Hent nye >>" />&nbsp;<a href="kontaktpers.asp?menu=kund&kid=<%=kid%>" class=rmenu target="_blank">+ Opret ny</a> 
				                </td>
				                </tr>
				                <tr>
				                <td>
				                <input type="checkbox" name="FM_att_vis" id="FM_att_vis" value="on" <%=attCHK %>><b>Vis Att.</b>&nbsp;
			                <br /><select name="FM_att" id="FM_att" style="width:150px; font-size:9px;">
				                <%
				                strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& kid &" ORDER BY navn"
				                oRec2.open strSQLkpersCount, oConn, 3
                					
					                while not oRec2.EOF
						                if cint(strAtt) = cint(oRec2("id")) then
						                attSel = "SELECTED"
						                else
						                attSel = ""
						                end if%>
						                <option value="<%=oRec2("id")%>" <%=attSel%>><%=oRec2("navn")%></option>
						                <%
					                oRec2.movenext
					                wend
					                oRec2.close
					                %>
				                <%if strAtt = "991" then
				                selth1 = "SELECTED"
				                else
				                selth1 = ""
				                end if%>	
				                <option value="991" <%=selth1%>>Den økonomi ansvarlige</option>
				                <%if strAtt = "992" then
				                selth2 = "SELECTED"
				                else
				                selth2 = ""
				                end if%>	
				                <option value="992" <%=selth2%>>Regnskabs afd.</option>
				                <%if strAtt = "993" then
				                selth3 = "SELECTED"
				                else
				                selth3 = ""
				                end if%>	
				                <option value="993" <%=selth3%>>Administrationen</option>
				                </select>
				                </td>
				                </tr>
		                </table>
		              
                  <!-- Modtager slut -->
   
	                </td><td valign=top width="50%">
	
	
	
	<!--- Afsender af faktura --->
	<table cellspacing="0" cellpadding="5" border="0" width="100%" bgcolor="#FFFFFF">
	<tr>
	    <td style="border:1px #ffffff solid;" bgcolor="#8caae6" class=alt><b>Afsender </b></td>
	</tr>
	<tr>
	    <td >
	
		<%
		afskid = 0
		
		strSQL = "SELECT adresse, postnr, city, land, telefon, fax, email, regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid FROM kunder WHERE useasfak = 1"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
			yourFax = oRec("fax")
		end if
		oRec.close
		
		
		%>
		  <div id="DIV1" style="width:300px; height:100px; padding:4px; border:1px #8CAAE6 solid; overflow:auto;">
                     
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td><%=yourNavn%></td>
		</tr>
		<tr>
			<td><%=yourAdr%></td>
		</tr>
		<tr>
			<td><%=yourPostnr%>&nbsp;&nbsp;<%=yourCity%></td>
		</tr>
		<tr>
			<td><%=yourLand%></td>
		</tr>
		</table>
		</div>
		
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
		    <td>
		    <br />Vor ref.: <select id="FM_vorref" name="FM_vorref" style="width:250px; font-size:10px;">
		    <option value="-">Ingen</option>
		    <%strSQlvorref = "SELECT mnavn, mid FROM medarbejdere WHERE mansat <> 2 AND mansat <> 3 ORDER BY mnavn "
		    oRec.open strSQlvorref, oConn, 3 
		    while not oRec.EOF  
		    
		    if (oRec("mid") = cint(session("mid")) AND func <> "red") OR (vorref = oRec("mnavn") AND func = "red") then
		    vfSEL = "SELECTED"
		    else
		    vfSEL = ""
		    end if  
		    
		    if cint(jobans1) = oRec("mid") then
		    vrjobans1 = " - (jobansvarlig)"
		    else
		    vrjobans1 = ""
		    end if
		    
		    
		    if cint(jobans2) = oRec("mid") then
		    vrjobans2 = " - (jobejer)"
		    else
		    vrjobans2 = ""
		    end if
		    %>
		    
		    
		    <option value="<%=oRec("mnavn") %>" <%=vfSEL %>><%=oRec("mnavn") %> <%=vrjobans1 %> <%=vrjobans2 %></option>
		    <%
		    oRec.movenext
		    wend
		    oRec.close%>
		    </select>
		    </td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_ytlf_vis" value="on" <%=visAfsTlfCHK %>> Vis Tlf: <%=yourTlf%></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_yfax_vis" value="on" <%=visAfsFaxCHK %>> Vis Fax: <%=yourFax%></td>
		</tr>
		<tr>
		    <td><input type="checkbox" name="FM_yemail_vis" value="on" <%=visafsemailCHK %>> Vis Email: <%=yourEmail%></td>
		</tr>
		</table>
		
		
		<br /><br />
		<div id="DIV2" style="width:300px; height:50px; padding:4px; border:1px #8CAAE6 solid; overflow:auto;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td><b>Bank information:</b></td>
		</tr>
		<tr>
			<td><%=yourBank%></td>
		</tr>
		<tr>
			<td>Reg. og Kontonr: <%=yourRegnr%> - <%=yourKontonr%></td>
		</tr>
		</table>
		</div>
		
			<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td><input type="checkbox" name="FM_yswift_vis" value="on" <%=visAfsSwiftCHK %>> Vis Swift: <%=yourSwift%></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_yiban_vis" value="on" <%=visAfsIbanCHK %>> Vis Iban: <%=yourIban%></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_ycvr_vis" value="on" <%=visAfsCvrCHK %>> Vis CVR nr.: <%=yourCVR%></td>
		</tr>
		</table>
		
		</td>
		</tr>
		
	
	</table>
	
	</td></tr>
	</table>
	</div>
	
	<%if jobid <> 0 then
	nst = "jobbesk"
	else
	nst = "aktdiv"
	end if %>
	
	    <div id=modtagdiv_2 style="position:absolute; visibility:hidden; display:none; top:496px; width:720px; left:5px; border:0px #8cAAe6 solid;">
        <table width=100% cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('fidiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('<%=nst %>')" class=vmenu>Næste >></a></td></tr></table>
		 </div>

	
	
	<!-- afsender slut -->
	
	
	
	        <%'** Fakdato MSG ****'
            itop = 140
            ileft = 455
            iwdt = 250
            idsp = "none"
            ivzb = "hidden"
            iId = "fakdatoinfo"
            call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
            %>
            Hvis du ændrer denne dato følger fakturadatoen ikke længere det valgte datointerval og timer kan gå tabt, eller blive faktureret dobbelt.
            <br /><br />
            Benyt evt. <b>Labeldato istedet</b>
		    <br /><br />Faktura (system) datoen der afgør hvor <b>skæringsdatoen</b> mellem timer på denne faktura og næste faktura skal ligge. 
		    <br />&nbsp;
	            </td></tr></table>
	            </div>
	        
            <input id="showfakmsg" value="0" type="hidden" />
	       
	
	
	
	
	<!-- Faktura indstillinger -->
	<!--<script type="text/javascript" src="../inc/jquery/jquery.erp.js"></script>-->
	<div id="fidiv" style="position:absolute; visibility:visible; top:105px; width:720px; left:5px; border:1px #8cAAE6 solid;">
    <table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
	<tr>
	    <td colspan=2 bgcolor="#8caae6" class=alt style="border:1px #ffffff solid; padding:5px 0px 5px 5px;"><b>Faktura indstilliger</b>
	    
	    &nbsp;&nbsp;(Type: <%
		'** Faktype ***
		select case intType
		case "0"
		strFaktypeNavn = "Faktura"
		case "1"
		strFaktypeNavn = "Kreditnota"
		case "2"
		strFaktypeNavn = "Rykker"
		case else
		strFaktypeNavn = "Faktura"
		end select
		%>
			
			<%=strFaktypeNavn %>)
	    
	    </td>
	 </tr>
	
	<tr>
		<td style="width:140px; padding:2px 5px 2px 5px;">
		
			<% 
	        if len(lastFakdato) <> 0 then
	        lastFakdato = lastFakdato
	        else
	        lastFakdato = "2001/1/1"
	        end if
	
	
		'** Periodeinterval brugt.
		'** Start dato
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		stdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&dagparset
		showStDato = dagparset&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
		
		
		'** Slut dato
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutdato = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&dagparset
		slutDagparset = dagparset
		showSlutDato = dagparset&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_aar") 
		useLastFakdato = 0
	
	    
	    'Response.Write lastFakdato
	    'Response.end
	    
	    '**** Hvis startdato ligger før sidste fakdato sættes stdato = sidste fakdato ****
	    if (cdate(showStDato) > cdate(lastFakdato) OR intType = 1) then 'kreditnota 'AND cint(usedt_ival) = 1 
		    stdatoKri = stdato
		    useLastFakdato = 0
	    else
		    stdatoKri_temp = dateadd("d", 1, lastFakdato)
		    stdatoKri = year(stdatoKri_temp)&"/"&month(stdatoKri_temp)&"/"&day(stdatoKri_temp)
		    showStDato = stdatoKri_temp 'lastFakdato
		    useLastFakdato = 1
	    end if
	    '******************************************************
	
	   
	
	   
	    %>  
	        <!-- jquery val for materiale fun --->
            <input id="jq_stdato" value="<%=stdatoKri%>" type="hidden" />
            <input id="jq_sldato" value="<%=slutdato%>" type="hidden" />
            <input type="hidden" id="jq_func" value="<%=func%>">
            <input type="hidden" id="jq_id" value="<%=id%>">
            <input type="hidden" id="jq_inttype" value="<%=intType%>">
		
		<b>Faktura dato:</b> (systemdato)<br />
				
		
		</td><td valign=top style="width:520px; padding:2px 5px 2px 5px;">
		
		<%'** Bruger afg. dato fra interval som faktura dato **'
		if func <> "red" then 'cint(usedt_ival) = 1 and
		strDag = datepart("d", showSlutDato, 2, 3)
		strMrd = datepart("m", showSlutDato, 2, 3)
		strAar = datepart("yyyy", showSlutDato, 2, 3)
		else
		strDag = datepart("d", strTdato, 2, 3)
		strMrd = datepart("m", strTdato, 2, 3)
		strAar = datepart("yyyy", strTdato, 2, 3)
		end if
		
		
		
		'Response.write "<b>" & replace(formatdatetime(strDag &"/"& strMrd &"/"& strAar, 2),"-",".") & "</b>"
		
		'if lastFakdato <> "2001/1/1" then
		'Response.Write "&nbsp;<font class=lillesort><i>(seneste faktura dato: " & replace(formatdatetime(lastFakdato, 2),"-",".") &")</i></font>" 
		'end if 
		%>
		<input type="text" name="FM_fakdato" id="FM_fakdato" value="<%=strDag%>.<%=strMrd%>.<%=strAar%>" style="display:none; margin-right:5px; width:70px;" />

        <input name="FM_start_dag" id="FM_start_dag" type="hidden" value="<%=strDag%>" />
        <input name="FM_start_mrd" id="FM_start_mrd" type="hidden" value="<%=strMrd%>" />
        <input name="FM_start_aar" id="FM_start_aar" type="hidden" value="<%=strAar%>" />
		
		
		<input type="hidden" name="FM_interval_slutdag" id="FM_interval_slutdag" value="<%=slutDagparset%>">
		<input type="hidden" name="FM_interval_slutmrd" id="FM_interval_slutmrd" value="<%=request("FM_slut_mrd")%>">
		<input type="hidden" name="FM_interval_slutaar" id="FM_interval_slutaar" value="<%=request("FM_slut_aar")%>">
		<!--<input type="hidden" name="showfakmsg" id="showfakmsg" value="0">-->
		
	    </td>
		</tr>
		
		<%
		
		if func <> "red" then
		
		    if request.Cookies("erp")("visperiode") <> "" then
    	         
	             if request.Cookies("erp")("visperiode") = "1" then
		         chkvp = "CHECKED"
		         else
    		     chkvp = ""
    		     end if
    		
		    else
    		
		         chkvp = ""
    		 
		    end if
		
		else
		     
		     if cint(visperiode) = 1 then
		     chkvp = "CHECKED"
		     else
    		 chkvp = ""
    		 end if
		
		end if
		
		
		%>
		
		
		<tr>
		<td valign=top style="padding:4px 5px 2px 5px;"><input id="Checkbox1" name="FM_visperiode" value="1" type="checkbox" <%=chkvp %> /> Vis <b>periode</b> på faktura:
		
		 <%if func <> "red" AND useLastFakdato = 1 then %>
        
	
		
		  <%
		  
		    uTxt = "Periode <b>start-dato</b> er korrigeret til dagen efter <br>seneste faktura dato på dette job."
	        uWdt = 250
	        call infoUnisportAB(uWdt, uTxt, 40, 430)
		  
		 
        end if %>
		
		
	    </td>
		
		
		
		<td valign=top style="padding:4px 5px 2px 5px;">
            
		<%if func <> "red" then
		istDato = showStDato
		istSlutDato = showSlutDato 
		else
		istDato = istDato
		istSlutDato = istDato2
		end if %>

		<!--#include file="inc/erp_istdato_inc.asp"--> til <!--#include file="inc/erp_istdato2_inc.asp"-->
		</td></tr>
		
		
		<%
		if func <> "red" then
		labelDato = now
		else
		labelDato = labelDato
	    end if
		
		
		
		
		if func <> "red" then
		    
		    'if lto = "execon" OR lto = "immenso" OR lto = "intranet - local" then
		    'chkflabel = "CHECKED"
		    
		    'else
		
		    if request.Cookies("erp")("flabel") <> "" then
    	         
	             if request.Cookies("erp")("flabel") = "1" then
		         chkflabel = "CHECKED"
		         else
    		     chkflabel = ""
    		     end if
    		
		    else
    		
		         chkflabel = ""
    		 
		    end if
		    
		    
		    'end if
		
		else
		     
		     if cint(brugfakdatolabel) = 1 then
		     chkflabel = "CHECKED"
		     else
    		 chkflabel = ""
    		 end if
		
		end if %>
		
		<tr><td valign=top style="padding:2px 5px 2px 5px;"><input id="Checkbox2" name="FM_brugfakdatolabel" value="1" type="checkbox" <%=chkflabel %> /><b>Labeldato:</b>
		<br /> Brug labeldato som fakturadato på fakturaprint.
		
		</td><td valign=top style="padding:4px 5px 2px 5px;"><!--#include file="inc/erp_labeldato_inc.asp"--></td></tr>
		
		
		<% 
		 if func = "red" then
	    strB_dato = strB_dato
	    adddayVal = dateDiff("d",strDag&"/"&strMrd&"/"&strAar, strB_dato,2,2)
	    else
	    end if
	    
	    %>
	    
	    <tr><td valign=top height=20 style="padding:7px 5px 0px 5px;">
		<b>Forfaldsdato:</b> (Kredit)</td><td valign=top style="padding:2px 5px 0px 5px;">
	    <%
	    
	    
	    
	    'if level <> 1 then
	    'select case lto
	    'case "execon", "immenso"
	    'hideffdato = 1
	    '''disTxt = "DISABLED"
	    'case else
	    'hideffdato = 0
	    '''disTxt = ""
	    'end select
	    'else
	    'hideffdato = 0
	    ''''disTxt = ""
	    'end if
	    
	     
	    call betalingsbetDage(betbetint, hideffdato)
    	    
	        if Not InStr(strForfaldsdato, "-") then
	        strForfaldsdato = strDag & "-" & strMrd & "-" & strAar
	        end if
    	    
	        'Response.write "strForfaldsdato: "& strForfaldsdato
    	    
    	    
	        reformatfordate = split(strForfaldsdato, "-")
	    
	   	
	  	
	  	if hideffdato <> 1 then%>
	  	<input type="text" id="FM_forfaldsdato" name="FM_forfaldsdato" value="<%=reformatfordate(0)%>.<%=reformatfordate(1)%>.<%=reformatfordate(2)%>" style="margin-right:5px; width:70px;"/>
        <%else %>
        <input type="hidden" name="FM_forfaldsdato" value="<%=reformatfordate(0)%>.<%=reformatfordate(1)%>.<%=reformatfordate(2)%>" style="margin-right:5px; width:70px;"/>
        <%end if %>
        </td>
		</tr>
		
		<tr><td valign=top height=20 style="padding:7px 5px 2px 5px;">
		<b>Status:</b></td><td valign=top style="padding:2px 5px 2px 5px;">
				<input type="radio" name="FM_betalt" value="0" checked>Kladde&nbsp;&nbsp;&nbsp;
				<input type="radio" name="FM_betalt" value="1"><font color="yellowgreen"><b>Godkendt</b></font>
		 
		
		<input id="FM_type" name="FM_type" value="<%=intType%>" type="hidden" />
		
		</td>
		</tr>
		
		<tr>
		<td style="padding:4px 5px 0px 5px;"><b>Rabat:</b>
         </td>
		<td style="padding:4px 5px 0px 5px;">
           
           <%
								 if cint(visrabatkol) <> 0 then
								 rbtCHK = "CHECKED"
								 else
								 tbtCHK = ""
								 end if
								 %>
								 
                                <input id="FM_visrabatkol" name="FM_visrabatkol" type="checkbox" <%=rbtCHK %>>&nbsp;Vis rabat-kolonne på faktura
            
                 
           </td>
		</tr>
		<tr><td valign=top style="padding:6px 5px 0px 5px;">
		<b>Fakturerings valuta:</b>
		
		</td>
		<td style="padding:4px 5px 0px 5px;">
		
		<%
		
		select case jobfaktype
		case 0
		jftp = 0
		case 1
		jftp = 1
		case else
		jftp = 0
		end select
		
		
		
		call selectAllValuta(1, jftp) 
		
		%>
							
            
		</td></tr>
		<tr><td style="padding:4px 5px 0px 5px;"><b>Sprog:</b>
		
		</td>
		<td style="padding:2px 5px 0px 5px;">
							<select name="FM_sprog" style="width:70px;">
							<%strSQL = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " 
							oRec.open strSQL, oConn, 3
							while not oRec.EOF 
							 if oRec("id") = cint(sprog) then
							 sSEL = "SELECTED"
							 else
							 sSEL = ""
							 end if%>
							<option value="<%=oRec("id") %>" <%=sSEL %>><%=oRec("navn") %> </option>
							<%
							oRec.movenext
							wend
							oRec.close%>
							
							</select>
		</td></tr>
		
		<%
		if cint(momssats) = 25 then
		msSEL0 = ""
		msSEL25 = "SELECTED"
		else
		msSEL0 = "SELECTED"
		msSEL25 = ""
		end if %>
		
		
		<tr><td style="padding:2px 5px 4px 5px;"><b>Moms:</b></td>
		<td style="padding:2px 5px 4px 5px;">
            <select id="FM_momssats" name="FM_momssats" style="width:70px;">
                <option value="25" <%=msSEL25 %>>25%</option>
                <option value="0" <%=msSEL0 %>>0%</option>
            </select></td>
		</tr>
		
		<%if len(trim(sideskiftH)) <> 0 then 
		sideskiftH = sideskiftH
		else
		sideskiftH = 0
		end if%>
		
		<tr><td style="padding:2px 5px 4px 5px;"><b>Sideskift efter:</b></td>
		<td style="padding:2px 5px 4px 5px;">
            <select name="FM_sideskiftlinier" id="Select1">
            <%for l = 0 to 30  %>
            
            
            
            
            <%if cint(sideskiftlinier) = l then
            sideskiftHSEL = "SELECTED"
            else
            sideskiftHSEL = ""
            end if
                
                if l <> 0 then%>
                <option value="<%=l %>" <%=sideskiftHSEL %>><%=l %></option>
                <%else %>
                <option value="<%=l %>" <%=sideskiftHSEL %>><%=l %> (ingen sideskift)</option>
                <%end if %>
                
            <%next %>
            </select>
            linier
           </td>
		</tr>
		
    </table>
    </div>





<div id="kontodiv" style="position:absolute; visibility:visible; width:720px; display:; top:465px; left:5px; border:1px #8cAAE6 solid;">
 <table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#ffffff">
 <tr>
     <td bgcolor="#8caae6" colspan=2 class=alt style="border:1px #ffffff solid; padding:5px 0px 5px 5px;"><b>Konto posteringer</b> (Valgfrit)</td>
	</tr>
	<tr><td valign=top style="padding:12px 5px 2px 20px;"><b>Opret posteringer på konti oprettet på kontoplanen:</b><br />
	Posteringer tilføjes ved godkend, konti oprettes i under bøgføring.</td></tr>
	<tr>
	<td valign=top style="padding:12px 5px 2px 20px;"><b>Debitor konto:</b><br />
	
		<%
		if func = "red" then
		debKontouse = intKonto
		else
		debKontouse = kid
		end if
		
		select case lto
		case "execon", "immenso"
		disa = "DISABLED"
		case else
		disa = ""
		end select
		
		%>
	    <select name="FM_kundekonto" <%=disa %> style="width:250; font-size:9px; background-color:#ffffe1;">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
			strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			&" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			&" ORDER BY k.kontonr, k.navn"
			oRec.open strSQL, oConn, 3 
			fod = 0
			while not oRec.EOF
			
				if (debKontouse = oRec("kid") AND func <> "red") OR _
				(debKontouse = oRec("kontonr") AND func = "red") then
				selkon = "SELECTED"
				fod = 1
				else
				selkon = ""
				end if
		
			
			%>
			<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%> - <%=oRec("momskode") %></option>
			<%
			oRec.movenext
			Wend 
			oRec.close
		
		
		if fod = 0 then
		    
		    if request.Cookies("erp")("debkonto") <> "" then
		    debkontonr = request.Cookies("erp")("debkonto")
		    
		    strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			&" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			&" WHERE kontonr = " & debkontonr &" ORDER BY k.kontonr, k.navn"
			
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    debkontonavn = oRec("navn")
		    debmomskode = oRec("momskode")
		    end if
		    oRec.close
		    
		    %>
		    <option value="<%=debkontonr%>">(<%=debkontonr%>)&nbsp;&nbsp;<%=debkontonavn%> - <%=debmomskode %></option>
			
		    <%
		    end if 
		
		end if%>
		</select><br />
        &nbsp;
		
		
       
		</td>
		<td valign=top style="padding:12px 5px 2px 20px;">
		<b>Kreditor konto:</b> (Intern konto)<br />
		
		<%
		fok = 0
				
				strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			    &" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			    &" ORDER BY k.kontonr, k.navn"
				
				%>
				<select name="FM_modkonto" <%=disa %> style="width:250px; font-size:9px; background-color:pink;">
		        <option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
				<%
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
				if intModKonto = oRec("kid") then
				selkon = "SELECTED"
				fok = 1
				else
				selkon = ""
				end if
				%>
				<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%> - <%=oRec("momskode") %></option>
				<%
				oRec.movenext
				Wend 
				oRec.close
		
		
		if fok = 0 then
		    if request.Cookies("erp")("krekonto") <> "" then
		    krekontonr = request.Cookies("erp")("krekonto")
		    
		    strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			&" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			&" WHERE kontonr = " & krekontonr &" ORDER BY k.kontonr, k.navn"
		    
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    krekontonavn = oRec("navn")
		    kremomskode = oRec("momskode")
		    end if
		    oRec.close
		    %>
		    <option value="<%=krekontonr%>" SELECTED>(<%=krekontonr%>)&nbsp;&nbsp;<%=krekontonavn%> - <%=kremomskode %></option>
			
		    <%
		    end if 
		
		end if%>
		</select><br />
		
		
            &nbsp;
		</td>
		</tr>
		<%
		
		'if func <> "red" then
		
		'if request.Cookies("erp")("momskonto") <> "" then
	         
	    '     if request.Cookies("erp")("momskonto") = "2" then
		'     chk2 = "CHECKED"
		'     chk1 = ""
    '		 else
    '		 chk1 = "CHECKED"
	'	     chk2 = ""
    '		 end if
	'	
	'	else
	'	
	'	    chk1 = "CHECKED"
	'	    chk2 = "" 
	'	 
	'	end if
	'	
	'	else
		     
	'	     if cint(momskonto) = 2 then
	'	     chk2 = "CHECKED"
	'	     chk1 = ""
    '		 else
    '		 chk1 = "CHECKED"
	'	     chk2 = ""
    '		 end if
	'	
	'	end if
		
		
		%>
		
        <input id="FM_momskonto" name="FM_momskonto" value="0" type="hidden" />
		<!--
		<tr><td colspan=2 valign=top style="padding:5px 5px 2px 20px;">
		<b>Momsberegning</b><br />
		Vælg hvilken konto's momsprocent der skal benyttes til momsberegning på faktura og tilhørende postering.<br />
		
		</td></tr>
		<tr>
		<td valign=top style="padding:2px 5px 5px 20px;"><input id="Radio1" type="radio" name="FM_momskonto" value="1" <%=chk1 %> /> Benyt debitkonto </td>
		<td valign=top style="padding:2px 5px 5px 20px;"><input id="Radio1" type="radio" name="FM_momskonto" value="2" <%=chk2 %> /> Benyt kreditkonto </td>
		</tr>
		-->
		
		</table>
		</div>
		
		<div id=fidiv_2 style="position:absolute; visibility:visible; display:; top:610px; width:720px; left:5px; border:0px #8cAAe6 solid;">
        <table width=100% cellspacing=0 cellpadding=5 border=0><tr><td></td><td align=right><a href="#" onclick="showdiv('modtagdiv')" class=vmenu>Næste >></a></td></tr></table>
		</div>
	<!-- Faktura indstillinger SLUT -->
	
	
	
	
	
	
	
	
	
	
	<%if jobid <> 0 AND aftid = 0 then %>
	<!-- Job beskrivelse -->
	<div id=jobbesk style="position:absolute; visibility:hidden; display:none; top:105px; width:720px; left:5px; border:1px #8cAAE6 solid;">
    <table cellspacing="0" cellpadding="5" border="0" width=100% bgcolor="#FFFFFF">
	<%if (len(trim(strJobBesk)) <> 0 AND request.Cookies("erp")("huskjobbesk") <> "0" AND func <> "red") OR (len(trim(strJobBesk)) <> 0 AND func = "red") then
	visJb = "CHECKED"
	else
	visJb = ""
	end if
	
	if fastpris = 1 then
	jType = "Fastpris"
	else
	jType = "lbn. timer"
	end if
	%>
	<tr>
		<td valign=top>
		<input type="checkbox" name="FM_visikkejobnavn" id="FM_visikkejobnavn" value="1" <%=visikkejobnavnCHK%>>Vis <b>ikke</b> jobnavn (overskrift) på faktura<br />
	
		<input type="checkbox" name="FM_visjobbesk" id="FM_visjobbesk" value="1" <%=visJb%>>Vis jobbeskrivelse på faktura
	    
	<%if jobstatus = 1 then
		diab = "" 
		    
		    
		    if lto = "essens" then
		    chklukjob = 1
		    end if
		    
		    if chklukjob = 1 then
		    lkjobCHK = "CHECKED"
		    else
		    lkjobCHK = ""
		    end if
		    
		else
	    diab = "DISABLED"
	    end if 
	    %>
	    
	    
		
			<br />
			<input type="checkbox" name="FM_lukjob" id="FM_lukjob" <%=lkjobCHK %> value="1" <%=diab %>>Luk job ved faktura oprettelse
	    
	<br />
		
		
		<div style="position:relative; left:14px; top:10px; padding:10px; width:300px; border:1px #8CAAe6 solid; background-color:#EFf3FF;">
    
		<b><%=strJobnavn %> (<%=intjobnr %>)</b><br />
		Jobtype: <b><%=jType %></b><br />
		
		<%select case jobstatus
		case 1
		jobstatusTxt = "Aktiv"
		case 2
		jobstatusTxt = "Passiv"
		case 0
		jobstatusTxt = "Lukket"
		end select %>
		
		Status: <b><%=jobstatusTxt%></b><br />
		Budget / Pris: <b><%=formatnumber(intJobTpris, 2) &" "& valutaKode %></b><br />
		Fakturerbare timer forkalk: <b><%=formatnumber(intBudgettimer, 2) %></b> timer<br />
		
		<!--
		Faktura grundlag: 
		<select case jobfaktype
		case 0
		Response.Write "<b>Den tid der bruges pr. medarb.</b><br>"
		case 1
		Response.Write "<b>Den aktivitet der udføres</b><br>"
		end select >
		-->
		
		
		Periode: <b><%=formatdatetime(jobstdato, 2) %></b> til <b><%=formatdatetime(jobsldato, 2) %></b><br />
		Rekvnr: <b><%=rekvnr %></b><br />
		
		<%if cint(ski) = 1 then 
		skiTxt = "Ja"
		skiCHK = "CHECKED"
		else
		skiTxt = "Nej"
		skiCHK = ""
		end if%>
		
		<input id="FM_fak_ski" name="FM_fak_ski" type="checkbox" <%=skiCHK %> value="1" />  SKI aftale 
		
		<%if lto = "execon" OR lto = "immenso" OR lto = "intranet - local" then %>
		
		<%if cint(abo) = 1 then %>
		<br /><br />Dette er en <b>Lightpakke</b>
            <input id="Hidden1" name="FM_fak_abo" value="1" type="hidden" />
		<%end if %>
		
		<%if cint(ubv) = 1 then %>
		<br />Jobbet er omfattet af <b>Udbudsvagten</b>
		<input id="Hidden2" name="FM_fak_ubv" value="1" type="hidden" />
		<%end if %>
		
		<%end if %>
		
		</div>
		
		
		
	</td>
    </tr>
	<tr><td style="padding:20px 20px 20px 20px;"><b>Jobbeskrivelse:</b><br />
	
	                    <%
	                    content = strJobBesk
            			
			            
			            Set editorJ = New CuteEditor
            					
			            editorJ.ID = "FM_jobbesk"
			            editorJ.Text = content
			            editorJ.FilesPath = "CuteEditor_Files"
			            editorJ.AutoConfigure = "Minimal"
            			
			            editorJ.Width = 680
			            editorJ.Height = 280
			            editorJ.Draw()
		                %>
	
	
	
	
	</td></tr>
	</table>
	</div>
	
	<div id=jobbesk_2 style="position:absolute; visibility:hidden; display:none; top:636px; width:720px; left:5px; border:0px #8cAAe6 solid;">
        <table width=100% cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('modtagdiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('aktdiv')" class=vmenu>Næste >></a></td></tr></table>
    
    
    
    </div>
	
	<!-- Job / aftale beskrivelse SLUT -->
	 <%end if 
	 
	 
	 '*** Mat og Job log ***
	 call joblogdiv()
	 call matlogdiv()
	    
	    
	%> 
	    
	    
	<!--<div id="aktiviteter" style="position:absolute; left:10px; top:920px; width:500px; visibility:visible; z-index:100; background-color:#d6dff5;">-->
	<%
	
	call akttyper2009(4)
	'Response.Write "aty_sql_onfak: " & aty_sql_onfak  
	'Response.end
	
	'*** Udspecificering af aktiviteter på job **'
	if jobid <> 0 AND aftid = 0 then '** jobid <> 0 (Vises kun ved fakturaer på job) %>
	<!--#include file="inc/fak_job_inc_2007.asp"-->
	<!--#include file="inc/mat_inc_2008.asp"-->        
	        
	<%else
	'*** Fakturerer Aftale ***'
	%>
	<!--#include file="inc/fak_aft_inc_2007.asp"-->
	<%end if %>
	
	
	<div id=betdiv style="position:absolute; width:720px; visibility:hidden; display:none; border:1px #8caae6 solid; top:105px; padding:0px 0px 0px 0px; left:5px; background-color:#EFf3FF;">
    
	
    <br />
	&nbsp;&nbsp;<b>Betalingsbetingelser / Kommentar:</b>
	<table width=100% cellspacing=5 cellpadding=0 border=0>
	<tr>
		<td style="padding:0px 0px 0px 5px;" valign="top">
		
		                <%
	                    content = strKom
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_komm"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "Minimal"
            			
			            editorK.Width = 680
			            editorK.Height = 280
			            editorK.Draw()
		                %>
		
		<br>
		<input type="checkbox" name="FM_gembetbet" id="FM_gembetbet" value="1">Gem betalingsbetingelser og forfaldsdato som standard for <b>denne</b> kontakt.<br>
		
		<%if level = 0 then%>
		<input type="checkbox" name="FM_gembetbetalle" id="FM_gembetbetalle" value="1">Gem som standard betalingsbetingelser og forfaldsdato for <b>alle</b> kontakter.<br>
		<%end if%>
		<input type="hidden" name="FM_kundeid" id="FM_kundeid" value="<%=intKid%>">
		</td>
	</tr>
	
	</table>
	
	<table cellpadding=0 border=0 cellspacing=0 width=90%>
	<tr>
		<td valign="top" align=right>
		<input name="subm_on" id="subm_on" type="submit" value="Se faktura" />
         </td>
    </tr>
	</table>
	
	 <div id=betdiv_2 style="position:relative; visibility:hidden; display:none; top:60px; width:600px; left:5px; border:0px #8cAAe6 solid;">
        <table width=580 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< Forrige</a></td><td align=right>
            &nbsp;</td></tr></table>
    </div>
	
	
	<!-- Nedenstående Bruges af javascript --->
	<input type="hidden" name="FM_showalert" id="FM_showalert" value="0">
	
	
	</div><!-- Betalingsbet -->
	<br>
        &nbsp;
<!--
</div>-->




    <%if jobid <> 0 then %>
	<div id="faksubtotal" style="position:absolute; left:730px; top:264px; width:200px; z-index:2000; border:1px #8caae6 solid; background-color:#ffffff; padding:5px;">
    <b>Faktura total:</b>
	<table width=100% cellspacing=5 cellpadding=0 border=0>
	<tr>
	<td align="right">Beløb ialt:
		<!-- strBeloeb -->
		<%
		totalbelob = totalbelob + matSubTotalAll
		totalbelob_umoms = (matSubTotalAlluMoms/1) + (aktSubTotalAlluMoms/1)
		
		if len(totalbelob) <> 0 then
		thistotbel = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbel = formatnumber(0, 2)
		end if
		
		if len(totalbelob_umoms) <> 0 then
		'Response.Write "" & totalbelob_umoms
		'Response.end
		thistotbel_Umoms = SQLBlessDot(formatnumber(totalbelob_umoms, 2))
		else
		thistotbel_Umoms = formatnumber(0, 2)
		end if
		%>
		<input type="hidden" id="FM_beloeb" name="FM_beloeb" value="<%=thistotbel%>">
		<input type="hidden" id="FM_beloeb_umoms" name="FM_beloeb_umoms" value="<%=thistotbel_Umoms%>">
		<div style="position:relative; width:90px; border-bottom:2px #86B5E4 dashed; padding-right:3px;" align="right" name="divbelobtot" id="divbelobtot"><b><%=thistotbel &" "& valutakodeSEL%></b></div>
		<br /><br />
		<div style="position:relative; width:120px; border-bottom:0px #86B5E4 dashed; font-family:arial; color:#999999; font-size:10px; padding-right:3px;" align="right" name="divbelobtot_umoms" id="divbelobtot_umoms">Beløb uden moms:<br /> (<%=thistotbel_Umoms &" "& valutakodeSEL%>)</div>
		
		</td>
	</tr>
	</table>
	</div>
	<%end if %>						
									
		
            <input id="lto" value="<%=lto %>" type="hidden" />				
		</form>							
									
								
</div> <!--sidetop -->
 <!-- main -->





<%
'** Viser menu efter side er loadet færdig ***'
 '** Sætter default værdier til enheder **'
if lto = "dencker" AND func <> "red" then
Response.Write("<script>opd_akt_endhed('Pr. enhed','2');</script>") 
end if

'Response.Write("<script>showmenu();</script>")


'*** Opdater valuta på sumfelter **'
if jobid <> 0 AND valuta <> 1 then
'Response.Write("<script>showvalutaload();</script>")
'Response.Write("<script>opdatervalutAllelinier(1,0);</script>")
'Response.Write("<script>hidevalutaload();</script>")
end if
%>

<%end select%>
<%end if





%>
<!--#include file="../inc/regular/footer_inc.asp"-->


