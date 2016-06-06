
<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->




<%

if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")

case "FN_sogktak"

        if len(trim(request("jqsogval"))) <> 0 then
		sogVal = request("jqsogval")
        else
        sogVal = ""
        end if
   
	    
    if sogVal <> "" then
	    
        lastKid = 0
		kpersStr = ""
		strSQL = "SELECT kp.navn, kp.email, kp.titel, k.kid, k.kkundenavn FROM kontaktpers AS kp LEFT JOIN kunder AS k on (k.kid = kp.kundeid) WHERE kp.navn LIKE '"&sogVal&"%' OR kp.email LIKE '"&sogVal&"%' OR k.kkundenavn LIKE '"&sogVal&"%' ORDER BY k.kkundenavn, kp.navn " & kid
		'Response.Write strSQL 
		'Response.end
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF

        if lastKid <> oRec("kid") then
        kpersStr = kpersStr & "<br><br><b>"& oRec("kkundenavn") & "</b><br>"
        end if
        
        kpersStr = kpersStr & "<i>"& oRec("navn") & "</i> "
        
        if len(trim(oRec("titel"))) <> 0 then
        kpersStr = kpersStr & " ("& oRec("titel") &") "
        end if
        
        
        kpersStr = kpersStr & " <a href='mailto:"&oRec("email")&"&subject=Faktura: "& varFaknr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
        

        lastKid = oRec("kid")

        oRec.movenext
        wend
        oRec.close


     '*** ÆØÅ **'
    call jq_format(kpersStr)
    kpersStr = jq_formatTxt
    response.write kpersStr

    end if


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
                            aftid = request("jq_aftid")
                            id = request("jq_id")
                            kuneks = request("jq_kuneks")
                            ingper = request("jq_ignper")
                            jq_intType = request("jq_inttype")
                            jq_fastpris = request("jq_fastpris")
                            
                            valutaKursFak = request("jq_valutaValkurs")
                            valutaLabelFak = request("jq_valutafaklabel")
                            'Response.Write func
                            'Response.end
                            matSubTotalAlluMoms = 0
                            matSubTotalAll = 0
                           
                                        'if session("mid") = 1 then
                                        'Response.write "jobid "& jobid & "<br>"
                                        'end if

                                    if len(trim(jobid)) <> 0 then
                                        jobid = jobid
                                    else
                                        jobid = 0
                                    end if

                                    if len(trim(aftid)) <> 0 then
                                        aftid = aftid
                                    else
                                        aftid = -1
                                    end if
                           
                            %>                            
                            <table width=99% cellspacing="0" cellpadding="1" border="0">
	                       
	                        <tr bgcolor="orange">
	                        <%

                            call matFelter
                            
                             '** Allerede fakturarede materialer 
	                        m = 0
	                        isMatWrt = " matid <> -1 "
	                        lastgrpnavn = ""
                        	
	                        if func = "red" then
                        	
	                        
	                        strSQL = "SELECT fms.matid, matnavn, matantal, matenhedspris AS matsalgspris, "_
	                        &" matenhed, matvarenr, ikkemoms, fms.valuta AS valutaid, fms.kurs, v.valutakode, matrabat, "_
	                        &" matfrb_mid AS mfusrid, matfrb_id AS mfid, matgrp, mgp.navn AS matgrpnavn, matbeloeb, fms_aktid AS mataktid FROM fak_mat_spec fms "_
	                        &" LEFT JOIN valutaer v ON (v.id = fms.valuta) "_
	                        &" LEFT JOIN materiale_grp AS mgp ON (mgp.id = fms.matgrp) "_
	                        &" WHERE matfakid = "& id &" ORDER BY mgp.navn, fms_aktid, matnavn"
	                        
                            'if session("mid") = 1 then
	                        'Response.Write strSQL
	                        'Response.flush
                             'end if
	                        'Response.end
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
                            
                              mataktid = oRec("mataktid")

                            matantal = oRec("matantal")
                            
                            
                            call materialer(1)
                            
                            isMatWrt = isMatWrt & " AND matid <> " & oRec("mfid")
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
                        	
                        	
	                        
                            end if
                            
                            

                            '** Ikke fakturerede Materialer / Eller v. ny faktura = alle 
	                        strSQL = "SELECT mf.matid, mf.matnavn, sum(mf.matantal) AS matantal, mf.matsalgspris, mf.matenhed, mf.matkobspris, "_
	                        &" mf.matvarenr, mf.valuta, v.valutakode, v.kurs, mf.erfak, mf.id AS mfid, "_
	                        &" mf.usrid AS mfusrid, intkode, matgrp, mgp.navn AS matgrpnavn, mf.ava, mf.aktid AS mataktid "_
	                        &" FROM materiale_forbrug mf "_
	                        &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	                        &" LEFT JOIN materiale_grp AS mgp ON (mgp.id = matgrp) "
                            
                            if jobid <> 0 then
	                        strSQL = strSQL &" WHERE (mf.jobid = "& jobid & ") "  
	                        else
                            strSQL = strSQL &" WHERE (mf.serviceaft = "& aftid &") "
                            end if

	                        if kuneks <> 0 then 'kun eksterne
	                        strSQL = strSQL &" AND mf.intkode = 2 " 
	                        end if
	                        
	                        if ingper <> 1 then
	                        strSQL = strSQL &" AND forbrugsdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"'"
	                        end if
	                        
	                        '***hent kun dem der ikke allerede er faktureret ***'
	                        strSQL = strSQL &" AND erfak = 0 " 
	                        
	                        strSQL = strSQL &" AND ("& isMatWrt &") GROUP BY mf.matid, mf.matsalgspris, mf.matnavn ORDER BY mgp.navn, mf.aktid, mf.matnavn"
	                        
                            'if session("mid") = 1 then
	                        'Response.Write strSQL
	                        'Response.flush
                             'end if
	                        
	                        oRec.open strSQL, oConn, 3
                            im = 0
                            while not oRec.EOF
                            
                            
                            '*** ÆØÅ **'
                            matnavn = oRec("matnavn")
                            call jq_format(matnavn)
                            matnavn = jq_formatTxt 
                            
                            
                            matAva = oRec("ava")
                   
                               
                            intMatRabat = intRabat
                            valutaMatId = oRec("valuta") 
                            valutaMatKode = oRec("valutakode") 
                            valutakodeSEL = valutaLabelFak 'valutaMatKode
                            valutaMatKurs = oRec("kurs")

                            mataktid = oRec("mataktid")

                            matantal = oRec("matantal")
                            
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
                  useK = 1
                 
                 if len(trim(request.Form("cust"))) <> 0 then
                 kid = request.Form("cust")
                 else
                 kid = 0
                 end if
                 
                  strSQLkpersAdr = "SELECT kid, kkundenavn AS navn, adresse, postnr, city AS town, "_
                  &" land, ean, cvr FROM kunder WHERE kid = "& kid
		   
            else
                 

                 useK = 2

                 if len(trim(request("att_val"))) <> 0 then
                 att = request("att_val")
                 else
                 att = 0
                 end if
          
            strSQLkpersAdr = "SELECT kp.id, kp.navn, kp.adresse, kp.afdeling, kp.postnr, kp.town, kp.land, kp.kpcvr AS cvr, kp.kpean AS ean, k.cvr AS cvrAlt "_
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
			                
			                select case lto
                            case "synergi1"
                                cvrTxt = "VAT"
                                cvrTxt2 = ""
                            case "epi_no", "epi_sta"
                                 cvrTxt = "Organisasjonsnr"
                                 cvrTxt2 = " MVA"
                            case else
                                 cvrTxt = "CVR"
                                cvrTxt2 = ""
                            end select


                            if cint(vis_cvr) = 1 then


                                if cint(useK) = 1 then 'kontakt

                                if isNull(oRec("cvr")) <> true AND len(trim(oRec("cvr"))) > 0 then
                                cvr = trim("<br>"& cvrTxt &": "& oRec("cvr") & cvrTxt2)
                                else
			                    cvr = "<br>"& cvrTxt &":"
                                end if

                                else

                                'if cint(useK) = 2 then 'kpers
                
                                    'if cvr = "<br>"& cvrTxt &":" then
                    
                    	                
                                          if isNull(oRec("cvr")) <> true AND len(trim(oRec("cvr"))) > 0 then 'der er oprettet CVR nr på kontaktperson
                                          cvr = trim("<br>"& cvrTxt &": "& oRec("cvr") & cvrTxt2)
                                          
                                          else

                                                if isNull(oRec("cvrAlt")) <> true AND len(trim(oRec("cvrAlt"))) > 0 then
                                                cvr = "<br>"& cvrTxt &": "& oRec("cvrAlt") & cvrTxt2
                                                else
			                                    cvr = "<br>"& cvrTxt &":"
                                                end if
                    
                                           end if
                
                                    'end if

                                 'end if

                                end if

			            
                            else
			                cvr = ""
			                end if
                                   
            
           
            strAdr = strAdr & oRec("navn")  
            
              if cint(useK) = 2 then 'kpers
                
                if isNull(oRec("afdeling")) <> true AND len(trim(oRec("afdeling"))) <> 0 then
                strAdr = strAdr & " - " & oRec("afdeling")
                end if 

             end if
            
            strAdr = strAdr & "<br>"& oRec("adresse") & "<br>"& oRec("postnr") & " " &oRec("town") & land & ean & cvr


            end if
            oRec.close
            
             '*** ÆØÅ **'
            call jq_format(strAdr)
            strAdr = jq_formatTxt


           Response.Write strAdr
        
 case "FN_getCustDesc_land"
            
          
            
            
         

                 if len(trim(request("att_val"))) <> 0 then
                 att = request("att_val")
                 else
                 att = 0
                 end if
          
            strSQLkpersAdr = "SELECT kp.land "_
            &" FROM kontaktpers AS kp WHERE kp.id = "& att
		   
          
            'Response.write strSQLkpersAdr
            'Response.end
             
            
            oRec.open strSQLkpersAdr, oConn, 3
            if not oRec.EOF then
                            
            strland = oRec("land")
			                
            end if
            oRec.close
            
             '*** ÆØÅ **'
            call jq_format(strland)
            strland = jq_formatTxt


           Response.Write strland     




case "FN_bankkonto"

        
        kidSel = request("kidsel")

        'Response.write "<option>"& kidSel &"</option>"
        'Response.end

        strSQLKonto = "SELECT kid, regnr, kontonr, regnr_b, kontonr_b, regnr_c, kontonr_c FROM kunder WHERE kid = "& kidSel 'Selskab, licensejer eller datter selskab **'
		oRec.open strSQLKonto, oConn, 3 
		if not oRec.EOF then
            

             strBankkonto = "<option value='0'>"& oRec("regnr") &" "& oRec("kontonr") &"</option>"
             strBankkonto = strBankkonto & "<option value='1'>"& oRec("regnr_b") &" "& oRec("kontonr_b") &"</option>" 
             strBankkonto = strBankkonto & "<option value='2'>"& oRec("regnr_c") &" "& oRec("kontonr_c") &"</option>"
           
            
            
        end if
        oRec.close  

            
            Response.write strBankkonto

End select
Response.End
end if
%>



<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<!--#include file="../inc/regular/erp_func.asp"-->
<!--include file="inc/dato2.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->






<%

    '** Bruges til PDF visning ***
    if len(request("nosession")) <> 0 then
    nosession = request("nosession")
    else
    nosession = 0
    end if




	if len(session("user")) = 0 AND cint(nosession) = 0 then
	%>
    <!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>


	<%

    level = session("rettigheder")

    func = request("func")

    if func = "dbopr" then
	id = 0
	else
	
	
	    if len(trim(request("id"))) <> 0 then
        id = request("id")
        else
        id = 0
        end if

    end if

	if len(trim(request("rdir"))) <> 0 then
	rdir = request("rdir")
    else
    rdir = ""
    end if


     if len(trim(request("FM_kunde"))) <> 0 then
    kid = request("FM_kunde")
  
        if kid <> 0 then
        kidSQL = " jobknr = "& kid
        aftKidSQL = " kundeid = "& kid 
        else
        kidSQL = " jobknr <> "& kid 
        aftKidSQL = " kundeid <> "& kid 
        end if 
       
    else
        if len(trim(request.Cookies("erp")("kid"))) <> 0 then
         kid = request.cookies("erp")("kid")
        kidSQL = " jobknr = "& request.Cookies("erp")("kid")
        aftKidSQL = " kundeid = "& request.Cookies("erp")("kid") 
        else
        kid = 0
        kidSQL = " jobknr <> "& kid 
        aftKidSQL = " kundeid <> "& kid 
        end if
    end if
    
    Response.Cookies("erp")("kid") = kid

	
	if len(request("FM_job")) <> 0 then
	jobid = request("FM_job")
	else
	    if len(request.cookies("erp")("jid")) <> 0 then
	    jobid = request.cookies("erp")("jid")
	    else
	    jobid = 0
	    end if
	end if
	
    response.cookies("erp")("jid") = jobid
	
	if len(request("FM_aftale")) <> 0 then
	aftaleid = request("FM_aftale")
	else
	    if len(request.cookies("erp")("aid")) <> 0 then
	    aftaleid = request.cookies("erp")("aid")
	    else
	    aftaleid = 0
	    end if
	end if

    aftid = aftaleid

    response.cookies("erp")("aid") = aftaleid

	
	if request("reset") = 1 then
	reset = 1 'request("reset")
	else
	reset = 0
	end if
	
	

         if cint(reset) = 1 then 
		        
		        if len(request("fid")) <> 0 then
		        fid = request("fid")
		        else
		        fid = 0
		        end if
		        
		  
	   end if
          

        if len(trim(request("visfaktura"))) <> 0 then
        visfaktura = request("visfaktura")
        else
        visfaktura = 0
        end if

        
        
        if len(trim(request("visminihistorik"))) <> 0 then
        visminihistorik = request("visminihistorik")
        else
        visminihistorik = 0
        end if

        if len(trim(request("visjobogaftaler"))) <> 0 then
        visjobogaftaler = request("visjobogaftaler")
        else
        visjobogaftaler = 0
        end if



        if request("formsubmitted") = "1" OR request("fromfakhist") = "1" OR rdir = "tilfak" then

            if len(trim(request("FM_jobonoff"))) <> 0 then
            vislukkedejob = 1
            else
            vislukkedejob = 0
            end if

        else

            if request.Cookies("erp")("jobonoff") <> "" then
            vislukkedejob = request.Cookies("erp")("jobonoff")
            else    
            vislukkedejob = 0
            end if

        end if


           if cint(vislukkedejob) = 1 then
		   jobonoffCHK = "CHECKED"
		   else
		   jobonoffCHK = ""
		   end if 

        response.Cookies("erp")("jobonoff") = vislukkedejob

          

        if request("formsubmitted") = "2" then

            if len(request("FM_igdato")) <> 0 then
	            igDato = 1
	        else
                igDato = 0
	        end if

        else

              if request.cookies("erpfak")("igdato") <> "" then
                igDato = request.cookies("erpfak")("igdato")
              else
	            igDato = 0
	          end if

        end if


                if cint(igDato) = 1 then
	            chkigDato = "CHECKED"
                else
                chkigDato = ""
                end if

        response.cookies("erpfak")("igdato") = igDato


        if request("formsubmitted") = "1" then

            if len(request("FM_sog")) <> 0 AND request("FM_sog") <> "Søg.." then
            sogKri = request("FM_sog")
            kSQLkri = "AND kkundenavn LIKE '"& sogKri &"%' OR kkundenr = '"& sogKri &"'"

            else

                sogKri = "Søg.."
                kSQLkri = "AND kid <> 0"   

            end if

        else

            if request.cookies("erpfak")("sog") <> "" AND request.cookies("erpfak")("sog") <> "Søg.." then
            sogKri = request.cookies("erpfak")("sog")
            kSQLkri = "AND kkundenavn LIKE '"& sogKri &"%' OR kkundenr = '"& sogKri &"'"
            else
            sogKri = "Søg.."
            kSQLkri = "AND kid <> 0"
            end if

        end if

        response.cookies("erpfak")("sog") = sogKri
        response.cookies("erpfak").expires = date + 1

		       
	    thisfile = "erp_oprfak_fs"
	    menu = "erp"
	    

        func = request("func")

        print = request("print")

        if len(trim(request("media"))) <> 0 then
        media = request("media")
        else 
        media = "n" 'request("print")
        end if


            if len(trim(request("vans"))) <> 0 then
	        vans = request("vans")
            else
            vans = 0
            end if





        select case media
        case "j", "pdf", "print"
        visaspopup = 1 
        case else
                if len(trim(request("nomenu"))) <> 0 then
                visaspopup = 1
                else
                visaspopup = 0
                end if
        end select
	    


        dim content
        a = 0
        dim aval
        redim aval(a)

        
	 
	    %>
        
<!--
        Variable:  id:<%=id %>, kid:<%=kid %>, jobid:<%=jobid %>, aftaleid:<%=aftaleid %>, visfaktura:<%=visfaktura %>, visjobogaftaler:<%=visjobogaftaler %>, visminihistorik:<%=visminihistorik %>, func:<%=func %><br />
      
	  -->
	
	<%if cint(visaspopup) = 0 then %>
	            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

                <SCRIPT language=javascript src="inc/fak_func_2007.js"></script>


	            
	
            
               
    
                <%call menu_2014() %>
    
   


                <div id="faktura_vmenu" style="position:absolute; left:90px; width:280px; top:102px; border:0px #999999 solid;">
                        <!--#include file="erp_opr_faktura_kontakter.asp"-->
            
                        <%
                        '** Mini historik og faktura opretteles/rediger links
                        if cint(visjobogaftaler) = 1 then %>
                        <!--#include file="erp_opr_faktura_kontojob.asp"-->
                        <%end if %>

            
                        <%
                        '** Mini historik og faktura opretteles/rediger links
                        if cint(visminihistorik) = 1 then 
                
                            select case func 
                            case "slet"
                        
                                sletTop = 355
                                sletLft = 0

                                call erpslet
                
                            case "sletja"


                                call erpsletja


                           case "fortryd"

                                call erpfortryd
                
                        
                             case else
                                    
                                    '** Opdaterer faktura indstillinger  **''
                                    if func = "opdprefak" then 
	                        
                                        call setFakPreInd()
                                    end if
                            
                            %>
                            <!--#include file="erp_opr_faktura.asp"-->
                            <%end select %>
                        <%end if %>

        
            
                </div>


                <%if cint(visfaktura) = 1 then %>
                <div id="faktura_main" style="position:absolute; width:1000px; left:390px; top:82px;">
        	            <!--#include file="erp_fak.asp"-->
                </div>
                <%end if %>


     <%end if 'visaspopup = 0 %>


    <%if (func = "slet" OR func = "sletja") AND cint(visaspopup) = 1 then
                            
                            %><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%


                                
                                sletTop = 40
                                sletLft = 40
                               


                            select case func 
                            case "slet"
                    
                                call erpslet
                
                            case "sletja"


                                call erpsletja   

                            end select
     else %>



       <%if cint(visfaktura) = 2 then 
           
            if cint(visaspopup) = 0 then%>
            <div id="faktura_main" style="position:absolute; width:1000px; left:390px; top:102px;">
            <%end if 
           
            sprog = 1 	
	        strSQLFak = "SELECT f.sprog FROM fakturaer AS f WHERE fid = "& id

            'response.write strSQLFak
            'response.flush

            oRec6.open strSQLFak, oConn, 3
            if not oRec6.EOF then

                sprog = oRec6("sprog")

            end if
            oRec6.close
            %>

            
        	<!--#include file="erp_fak_godkendt_2007.asp"-->
            
            <%if cint(visaspopup) = 0 then%>
            </div>
            <%end if %>
        <%end if %>

    <%end if %>

    

   


	
	
	
	
	<%end if'session %>
	






