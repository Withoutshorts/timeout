

<%

    'id = request("id") '** Fakid
	showmoms = 0
    public strXML
'	
	'if len(request("jobid")) <> 0 then
	'jobid = request("jobid")
	'else
	'jobid = 0
	'end if
	
	'if len(request("aftid")) <> 0 then
	'aftid = request("aftid")
	'else
	'aftid = 0
	'end if
    
    'if lto = "essens" then
    'response.write "jobid:"& jobid & "id:"& id & "media:"& media
    'end if

    'response.write "thisfile: "& thisfile

    'Response.Write media
    'Response.flush
	
	select case media
	case "j", "pdf", "print" 

        leftAdd = 10
        
        
    
	     
        select case lto 
        case "synergi1", "xintranet - local"
        %>
        <!--#include file="../inc/regular/header_hvd_css3_html5_inc.asp"-->
        <%
        case else
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
        <%end select
	

     'Response.write "lto: " & lto & "<br>"

	case "xml"
	    
	    session.LCID = 1033
	    'Response.Charset = "utf-8"

        faktypeXML = 0
        strSQL = "SELECT faktype FROM fakturaer WHERE fid = "& id
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then
        faktypeXML = oRec("faktype")
        end if
        oRec.close
	    

        if cint(faktypeXML) <> 1 then 'faktura else kreditnota
        strXML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
        strXML = strXML &"<Invoice xmlns=""http://rep.oio.dk/ubl/xml/schemas/0p71/pie/"""_
        &" xmlns:com=""http://rep.oio.dk/ubl/xml/schemas/0p71/common/"""_
        &" xmlns:main=""http://rep.oio.dk/ubl/xml/schemas/0p71/maindoc/"""_
        &" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"""_
        &" xsi:schemaLocation=""http://rep.oio.dk/ubl/xml/schemas/0p71/pie/ "_
        &" http://rep.oio.dk/ubl/xml/schemas/0p71/pie/piestrict.xsd"">"
  
	    else
        strXML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
        strXML = strXML &"<Invoice xmlns=""http://rep.oio.dk/ubl/xml/schemas/0p71/pcm/"""_
        &" xmlns:com=""http://rep.oio.dk/ubl/xml/schemas/0p71/common/"""_
        &" xmlns:main=""http://rep.oio.dk/ubl/xml/schemas/0p71/maindoc/"">"
        end if



	case else
	
	  %>
            
       
        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    

        <script type="text/javascript">
	        $(document).ready(function () {




	            $("#oioxml").click(function () {

	                var xmlurl = $("#xmlurl").val()
	                var showvans = $("#showvans").val()
	                //alert(xmlurl)

	                function show_confirm() {
	                    var r = confirm("Ønsker du at afsende en OIO XML faktura til kunden via E-posthus (klik Ok), eller vil du blot gemme en ny kopi af OIO XML filen (klik Cancel)?");
	                    if (r == true) {
	                        window.open(xmlurl + "&vans=1", '_blank');

	                    }
	                    else {
	                        window.open(xmlurl + "&vans=0", '_blank');
	                    }
	                }

	                if (showvans == 1) {

	                    show_confirm()

	                } else {

	                    window.open(xmlurl + "&vans=0", '_blank');

	                }

	                //confirm("Ønsker du at afsende en OIO XML faktura til kunden via E-posthus, eller vil du blot gemme en ny kopi af OIO XML filen?")
	            });


	            //$("#bg_page2").css("top", 200)



	        });

        </script>


      <%
	    
	end select

    %>
    <script src="inc/erpfakgk_jav.js"></script>
    <!--#include file="inc/erp_fak_layout_inc_2007.asp"-->
	<%
	
	
	
	
	visjoblog = 0
	vismatlog = 0
    fakuraKolonneOverskrifterIsWrt = 0
	
	
    '** LUFT i toppen på side 2 /3 
	select case lto
	case "dencker", "tooltest", "outz", "intranet - local", "optionone"
	sidetoluft = 100
    case "synergi1"
    sidetoluft = 260
	case else
	sidetoluft = 20
	end select

	
	
	'*** Faktura info ****
	if jobid <> 0 then
	strSQL_sel = " jobid, jobnavn, jobnr, rekvnr, fastpris, "
	strSQL_lftJ = " job j ON (j.id = jobid) "

            select case lto
            case "nt", "intranet - local"
            strSQL_sel = strSQL_sel & "supplier_invoiceno, "
            
            end select

	else
	strSQL_sel = " s.navn AS aftnavn, s.aftalenr, "
	strSQL_lftJ = " serviceaft s ON (s.id = aftaleid) "
	end if
	
	strSQL = "SELECT f.editor, tidspunkt, fid, faknr, fakdato, b_dato, timer, beloeb, f.kommentar, f.betalt, "_
	&" tidspunkt, fakadr, att, f.faktype, konto, modkonto, "& strSQL_sel &" "_
	&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail,"_
	&" visafsswift, visafsiban, visafscvr, moms, enhedsang, f.varenr, jobbesk, "_
	&" visjoblog, visrabatkol, vismatlog, rabat, "_
	&" visjoblog_timepris, visjoblog_enheder, visjoblog_mnavn, "_
	&" visafsfax, betalt, f.valuta, f.kurs, v.valutakode AS valutaISO, f.sprog, "_
	&" f.istdato, momskonto, visperiode, istdato2, brugfakdatolabel, "_
	&" vorref, fak_ski, momssats, modtageradr, showmatasgrp, hidesumaktlinier, "_
	&" sideskiftlinier, f.subtotaltilmoms, labeldato, fak_abo, fak_ubv, visikkejobnavn, hidefasesum, "_
    &" hideantenh, medregnikkeioms, afsender, vis_jobbesk, kontonr_sel, usealtadr, fak_fomr"_
	&" FROM fakturaer f "_
	&" LEFT JOIN "& strSQL_lftJ &""_
	&" LEFT JOIN valutaer v ON (v.id = f.valuta) WHERE fid = "& id
	
	'Response.write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	
	fak_valuta = oRec("valuta")
	fak_kurs = oRec("kurs")
	
	valutaISO = oRec("valutaISO")
	kurs = oRec("kurs")
	
	strEditor = oRec("editor")
	fakTidspunkt = oRec("tidspunkt")
	intFaktype = oRec("faktype")
        
       
    intKundeid = oRec("fakadr")
       
    intAtt = oRec("att")
	
	'intKonto = oRec("konto")
	'intModKonto = oRec("modkonto")
	
	'if oRec("momskonto") <> 1 then
	'momsKonto = intModKonto
	'else
	'momsKonto = intKonto
	'end if
	
    intGodkendt = oRec("betalt")
	
	visjoblog = oRec("visjoblog")
	
	visjoblog_enheder = oRec("visjoblog_enheder")
	visjoblog_mnavn = oRec("visjoblog_mnavn")
	visjoblog_timepris = oRec("visjoblog_timepris")
	
	vismatlog = oRec("vismatlog")
	
	visrabatkol = oRec("visrabatkol")
	
	momssats = oRec("momssats")
	
	intFaktureretTimer = oRec("timer")
	intFaktureretBelob = oRec("beloeb")
	
	subtotalTilmoms = oRec("subtotaltilmoms")
	
	if intFaktype = 1 AND media <> "xml" then
	intFaktureretBelob = -(intFaktureretBelob)
	subtotalTilmoms = -(subtotalTilmoms)
	end if
	
   
	
	    if jobid <> 0 then
	    strKomm = oRec("kommentar")
	    strJobnavn = oRec("jobnavn")
	    intJobnr = oRec("jobnr")
	    strJobBesk = oRec("jobbesk")
    	buyersordId = intJobnr

	    if trim(len(oRec("rekvnr"))) <> 0 then
	    rekvnr = oRec("rekvnr")
        buyersordId = rekvnr
	    else
	    rekvnr = "0"
	    end if
    	

        fastpris = oRec("fastpris")


            select case lto
            case "nt", "intranet - local"
            supplier_invoiceno = oRec("supplier_invoiceno")
            end select


	    else
    	 

         'intJobnr = 0
	     rekvnr = "0"
	     strAftNavn = oRec("aftnavn")
	     strAftVarenr = oRec("varenr")
	     intAftnr = oRec("aftalenr")
         buyersordId = intAftnr
    	    
	        if cdate(oRec("fakdato")) > cdate("1/5/2007") then
	        strKomm = oRec("kommentar")
	        strJobBesk = oRec("jobbesk")
	        else
	        strJobBesk = oRec("kommentar")
	        strKomm = "Netto Kontant 8 dage."
	        end if
    	
        fastpris = 0

	    end if
	

    
	
	forfaldsdato = oRec("b_dato")
	varFaknr = oRec("faknr")
	istDato = oRec("istdato")
	istDato2 = oRec("istdato2")
	brugfakdatolabel = oRec("brugfakdatolabel")
	
	'** Label eller system dato ***'
	if cint(brugfakdatolabel) <> 1 then
	fakdato = oRec("fakdato") 
	else
	fakdato = oRec("labeldato") 'istDato2
	end if
	
    sysFakdato = oRec("fakdato")
	
	intVismodland = oRec("vismodland")
	intVismodatt = oRec("vismodatt")
	intVismodtlf = oRec("vismodtlf")
	intVismodcvr = oRec("vismodcvr")
	
	
	intVisafstlf = oRec("visafstlf")
	intVisafsemail = oRec("visafsemail")
	intVisafsswift = oRec("visafsswift")
	intVisafsiban = oRec("visafsiban")
	intVisafscvr = oRec("visafscvr")
	intVisafsfax = oRec("visafsfax")
	
	'if oRec("valuta") <> 1 then
	'intMoms = oRec("moms") * (100/kurs)
	'else
	intMoms = oRec("moms") 
	'end if
	
	if intFaktype = 1 AND media <> "xml" then
	intMoms = -(intMoms)
	end if
	
	select case oRec("enhedsang")
	case -1
	enhed = "" 'Ingen (skjul)
	case 1	
	enhed = erp_txt_019 '"Stk. pris"
    case 2
	enhed = erp_txt_020 '"Enhedspris"
	case else
	enhed = erp_txt_018 '"Timepris"
	end select
	
	enhedsangFak = oRec("enhedsang")
	
	intAftRabat = oRec("rabat")
	intStatus = oRec("betalt") 
	
    'sprog = oRec("sprog") 'Sættes i _fs filen, da den bestemmer XML lang settings
	
	visPeriode = oRec("visperiode")
	
	modtageradr = oRec("modtageradr")
	
	vorref = oRec("vorref")
	fak_ski = oRec("fak_ski")
	fak_abo = oRec("fak_abo")
	fak_ubv = oRec("fak_ubv")
	
	showmatasgrp = oRec("showmatasgrp")
	
	hidesumaktlinier = oRec("hidesumaktlinier")
	sideskiftlinier = oRec("sideskiftlinier") 

    '**Tvungen sideskift bruge i brødteskt **'
    'if cint(brpageUsed) = 1 then
    'sideskiftlinier = 31
    'end if
	
	visikkejobnavn = oRec("visikkejobnavn")
    hidefasesum = oRec("hidefasesum")

    hideantenh = oRec("hideantenh")

    medregnikkeioms = oRec("medregnikkeioms")
    afsender = oRec("afsender")

    vis_jobbesk = oRec("vis_jobbesk")

    kontonr_sel = oRec("kontonr_sel")

    usealtadr = oRec("usealtadr")

    fak_fomr = oRec("fak_fomr")
	
	if media = "xml" then
	
	session.LCID = 1033
	
	strXML = strXML & "<com:ID>"& varFaknr &"</com:ID>" 
	
	if month(fakdato) < 10 then
            monthFakd = "0"& month(fakdato)
            else
            monthFakd = month(fakdato)
            end if
            
            if day(fakdato) < 10 then
            dayFakd = "0"& day(fakdato)
            else
            dayFakd = day(fakdato)
            end if
            
            fakdatoXML = year(fakdato) &"-"& monthFakd &"-"& dayFakd
	
	strXML = strXML & "<com:IssueDate>"& fakdatoXML &"</com:IssueDate>" 
    
    if intFaktype = 0 then
    pieTyp = "PIE" 'Faktura
    else
    pieTyp = "PCM" 'Kreditnota
    end if
    
    strXML = strXML & "<com:TypeCode>"& pieTyp &"</com:TypeCode>" 
    strXML = strXML & "<main:InvoiceCurrencyCode>"& valutaISO &"</main:InvoiceCurrencyCode>" 
    
        if jobid <> 0 then
           
           if cint(visikkejobnavn) <> 1 then
           xmlNote = "Job: " & strJobnavn & "("& intJobnr &")"
           else
           xmlNote = ""
           end if
           
           if len(trim(strJobBesk)) <> 0 then
	       xmlNote = xmlNote & "<br>"&  strJobBesk
	       end if
		
	    else
	   
       intJobnr = intAftnr
	   xmlNote = strAftNavn & "(" & intAftnr & ")<br> Aftale nr.: "& strAftVarenr
	    
	    end if
    
    'call utf_format(xmlNote)
    'xmlNote = utf_formatTxt
          
    call htmlparseCSV(xmlNote)
    xmlNote = htmlparseCSVtxt
    
    xmlNote = replace(xmlNote, "''", "")
    
    strXML = strXML & "<com:Note><![CDATA["&EncodeUTF8(xmlNote)&"]]></com:Note>"
    
    
    end if
	
	
	end if
	oRec.close 
	
        

    	if media = "n" then
		bd = 1
		else
		bd = 0
		end if 
		

        '**************************************************************************************************************************************************************************************************
        '**** Bg på faktura print / PDF *******************************************************************************************************************************************************************
        '**************************************************************************************************************************************************************************************************

        globalLeft = 0 + leftAdd
       
		
		'*** sideskift ****'
		select case lto
        case "dencker"
        gblWdt = 620
                if media = "pdf" then%>
		
                <div id="fakturaside_bg" style="position:absolute; left:<%=globalLeft%>px; top:0px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; z-index:1; padding:0px 0px 0px 0px;">
                <img src="../ill/dencker/brevpapir_bg_850_2013.gif" />
                </div>
                <div id="fakturaside" style="position:relative; left:<%=globalLeft%>px; top:0px; width:640px; visibility:visible; border:<%=bd%>px #8caae6 solid; z-index:1000; padding:60px 0px 0px 80px;">
        
                <%
                else
                %>
                <div id="fakturaside" style="position:relative; left:<%=globalLeft%>px; top:0px; width:640px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
                <%
                end if 
        
        case "synergi1", "xintranet - local"
        gblWdt = 520
        
        if media = "n" then
        %>
         <div id="Div2" style="position:absolute; left:<%=globalLeft%>px; top:0px; width:<%=gblWdt+20%>px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
        <%
        else

        bg_d_topPDF = 85

           if media = "pdf" then
           firstbgDivL = -14
           else
           firstbgDivL = -8
           end if
     
          %>
          
            <div class="fakbg" id="Div3" style="position:absolute; top:<%=bg_d_topPDF%>px; width:758px; padding:0px; left:<%=firstbgDivL%>px; border:0px #000000 solid; z-index:-1; overflow:hidden;">
            <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/synergi/synergi_12.jpg" width="758" height="969" border="0" />
            </div>

             


             <%if sideskiftlinier >= 31 then 'Maks 5 sider
             
             s_slut = (sideskiftlinier - 30)

             for s = 1 To s_slut 

             
           

                if media = "pdf" then

                select case s
                case 1 'side 2
                bg_d_topPDF = 1230 '670 'ok
                bg_d_bgcol = "#FFFFFF"
                case 2 'side 3
                bg_d_topPDF = 2375 'ok
                bg_d_bgcol = "#FFFFFF"
                case 3 'side 4
                bg_d_topPDF = 3520
                bg_d_bgcol = "#FFFFFF"
                case 4 'side 5
                bg_d_topPDF = 4665
                bg_d_bgcol = "#FFFFFF"
                case 5 'side 6
                bg_d_topPDF = 5810
                bg_d_bgcol = "#FFFFFF"
                case 6 'side 7
                bg_d_topPDF = 6955
                bg_d_bgcol = "#FFFFFF"
                case 7 'side 8
                bg_d_topPDF = 8100
                bg_d_bgcol = "#FFFFFF"
                end select 


                %>
                <!-- 758 / 969 -->
               
               
                
               <div class="fakbg" id="fakturaside_<%=s %>" style="position:absolute; top:<%=bg_d_topPDF%>px; background-color:<%=bg_d_bgcol%>; width:758px; padding:0px; left:-14px; border:0px #000000 solid; z-index:<%=(2000+s*10)%>; padding:0px 0px 0px 0px; overflow:auto;">
             <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/synergi/synergi_12.jpg" width="758" height="969" border="0" style="top:0px;"/>
             </div>
             
             
                <%

                

                else
                
                select case s
                case 1
                bg_d_topPrint = 1280 'ok
                case 2
                bg_d_topPrint = 2425 'ok
                case 3
                bg_d_topPrint = 3570 'ok
                case 4
                bg_d_topPrint = 4715 'ok
                case 5
                bg_d_topPrint = 5860 'ok
                case 6
                bg_d_topPrint = 7005 'ok
                case 7
                bg_d_topPrint = 8150 'ok
                end select 
                %>
                
                <div id="Div5" style="position:absolute; top:<%=bg_d_topPrint%>px; height:969px; width:758px; padding:0px; left:-8px; border:0px #000000 solid; z-index:-1;">
             <img alt="" src="https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/synergi/synergi_12.jpg" width="758" height="969" border="0" />
             </div>
                <%
                end if%>


             <%next %>
              

             <%end if %>
            <div id="fakturaside_0" style="position:absolute; left:5px; top:0px; width:<%=gblWdt%>px; visibility:visible; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
        <%end if


        case "essens"
        gblWdt = 678
        %>
		<div id="fakturaside" style="position:relative; left:<%=globalLeft%>px; top:0px; width:698px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		

        <%
        case else
        gblWdt = 620
        %>
		
		<div id="fakturaside" style="position:relative; left:<%=globalLeft%>px; top:0px; width:640px; visibility:visible; background-color:#ffffff; border:0px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		
		
		<% 

        end select


        '**************************************************************************************************************************************************************************************************
        '**************************************************************************************************************************************************************************************************
        '**************************************************************************************************************************************************************************************************

	
	

	' *** Modtager ****
    if cdbl(usealtadr) <> 0 then
    strSQL = "SELECT kp.id AS kid, kp.navn AS kkundenavn, kp.afdeling, "_
    &" kp.adresse, kp.postnr, kp.town AS city, "_
    &" kp.land, kp.dirtlf AS telefon, kp.kpcvr AS cvr, kp.kpean AS ean, k.kkundenr, k.cvr AS cvrAlt "_
    &" FROM kontaktpers AS kp "_
    &" LEFT JOIN kunder AS k ON (k.kid = kp.kundeid) WHERE kp.id =" & usealtadr
    else
	strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr, ean FROM kunder WHERE Kid = " & intKundeid		
	end if

        'if session("mid") = 1 AND lto = "nt" then
        'response.write strSQL
        'response.flush
        'end if

    	oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			
			intKnr = oRec("kkundenr") 'bruges kun til beregning af FI nummer vises IKKE på faktura,
			strKnavn = oRec("kkundenavn")

             if cdbl(usealtadr) <> 0 then
                
                if isNull(oRec("afdeling")) <> true AND len(trim(oRec("afdeling"))) <> 0 then
                strKnavn = strKnavn & " - " & oRec("afdeling")
                end if 

             end if

			strKadr = oRec("adresse")
			strKpostnr = oRec("postnr")
			strBy = oRec("city")
			strLand = oRec("land")
			intKid = oRec("Kid")
			
            if isNull(oRec("cvr")) <> true then
            intCVR = oRec("cvr")
            else
			intCVR = ""
            end if

            if cdbl(usealtadr) <> 0 then
                
                if intCVR = "" then
                    
                    	
                    if isNull(oRec("cvrAlt")) <> true then
                    intCVR = oRec("cvrAlt")
                    else
			        intCVR = ""
                    end if
                    

                
                end if

             end if

            intTlf = oRec("telefon")
			ean = trim(oRec("ean"))
			
			if media = "xml" then
			'*** XML format kan kun sendes til oprindelig kunde **'
			
			session.LCID = 1033
			
			'call utf_format(strKnavn)
            'strKnavn = utf_formatTxt
            
          
          
          'Response.Write strKadr
          'Response.end
          
            
			
			call htmlparseCSV(strKadr)
            strKadr = htmlparseCSVtxt
          
            'call utf_format(strKadr)
            'strKadr = utf_formatTxt
          
            'call utf_format(strBy)
            'strBy = utf_formatTxt
			
			
		    strXML = strXML & "<com:BuyersReferenceID schemeID=""EAN"">"&ean&"</com:BuyersReferenceID>"
		    strXML = strXML & "<com:ReferencedOrder>"
            strXML = strXML & "<com:BuyersOrderID>"&buyersordId&"</com:BuyersOrderID>" 
            strXML = strXML & "<com:IssueDate>"&fakdatoXML&"</com:IssueDate>" 
            strXML = strXML & "</com:ReferencedOrder>"
		    strXML = strXML & "<com:BuyerParty>"
            strXML = strXML & "<com:ID schemeID=""CVR"">"&intCVR&"</com:ID>"
            strXML = strXML & "<com:PartyName>"
            strXML = strXML & "<com:Name><![CDATA["&EncodeUTF8(strKnavn)&"]]></com:Name>"
            strXML = strXML & "</com:PartyName>"
            strXML = strXML & "<com:Address>"
            strXML = strXML & "<com:ID>Fakturering</com:ID>" 
            strXML = strXML & "<com:Street><![CDATA["&EncodeUTF8(strKadr)&"]]></com:Street>"
           'strXML = strXML & "<com:HouseNumber>0</com:HouseNumber>" 
           strXML = strXML & "<com:CityName><![CDATA["&EncodeUTF8(strBy)&"]]></com:CityName>"
          strXML = strXML & "<com:PostalZone>"& strKpostnr &"</com:PostalZone>" 
          strXML = strXML &"<com:Country><com:Code listID=""ISO 3166-1"">DK</com:Code></com:Country>" 
          strXML = strXML & "</com:Address>"
          end if
         

			
		end if
		
		oRec.close
		
		
    
    
    finummer = ""
    
   
    select case lto
	case "dencker" , "intranet - local" , "synergi1"
	
	    select case lto 
	    case "dencker" 
	    kreditorFi = "80200552"
        case "synergi1", "intranet - local"
        kreditorFi = "80176856"
	    case else
	    kreditorFi = "00000000"
	    end select
    
    
    '*** Debitor id skal være 8 cifre **'
  
    lenintKnr = len(trim(intKnr))
    intKnr = replace(intKnr, " ", "")

    select case lenintKnr
    case 8 
    fiKnr = trim(intKnr)
    case else
    fiKnr = "00000000"
    end select
    
	finummer = finummer & fiKnr 
	
	
	'*** Fakturanr ID skal være 6 cifre **'
	lenvarFaknr = len(trim(varFaknr))
    select case lenvarFaknr
    case 6 
    fiFaknr = trim(varFaknr)
    case 5 
    fiFaknr = "0"&trim(varFaknr)
    case 4 
    fiFaknr = "00"&trim(varFaknr)
    case 3 
    fiFaknr = "000"&trim(varFaknr)
    case 2 
    fiFaknr = "0000"&trim(varFaknr)
    case 1 
    fiFaknr = "00000"&trim(varFaknr)
    
    case else
    fiFaknr = right(varFaknr,6)
    end select
	
	finummer = finummer & fiFaknr 
	
	
	
	'** Modulus tjek nummer ***'
	betTVsumCounter = 0
	for x = 1 to 14
	 
	 betTVsum = 0
	 betIndet = mid(finummer,x,1)
	 
     if len(trim(betIndet)) <> 0 AND lto <> "intranet - local" then

	 select case x
	 case 1,3,5,7,9,11,13
	 betIndet = (betIndet*1)
	 case else
	 betIndet = (betIndet*2)
	 end select
	 
     else

     betIndet = 0

     end if

	 'if session("mid") = 1 then
	 'Response.Write betIndet & ": "
	 'end if
        
	    for t = 1 to len(betIndet)
	    betTVsum = betTVsum + mid(betIndet,t,1)
        next

        betTVsumCounter = betTVsumCounter + betTVsum 
	    
        'if lto = "synergi1" AND session("mid") = 1 then
	    'Response.Write betTVsumCounter & "::" & betTVsum & "<br> "
        'end if
	    
	next
	

    
	
	modulus = 0
	modulus_division = (betTVsumCounter/10)
	
	komma = 0
	komma = instr(modulus_division, ",")
	
    
	'Response.Write "lenvarFaknr" & varFaknr & "<br><br>"
	
   
	
	if komma <> 0 then
	modulus_division = mid(modulus_division,1,komma-1)
	else
	modulus_division = modulus_division 
	end if
	
	
	

    'if session("mid") = 1 then
	'Response.Write "<br>betTVsumCounter:" & betTVsumCounter
	'Response.Write "<br>modulus_division:" & modulus_division
    'end if
	
	modulus = betTVsumCounter - (10 * modulus_division)
	kontrolciffer = (10 - modulus)
	
    if kontrolciffer = 10 then
    kontrolciffer = 0
    end if

    'if session("mid") = 1 then
	'Response.Write "<br>modulus:" & modulus
	'Response.Write "<br>kontrolciffer:" & kontrolciffer
	'end if

	finummer = "+71<" & finummer & kontrolciffer & "+"& kreditorFi & "<"
    
    end select
    
    
    kid = intKid'** Bruges ikke
		
	'*** Att ******
	select case intAtt
	case "0"
	showAtt = ""
	case "991"
	showAtt = "Den økonomi ansvarlige"
	case "992"
	showAtt = "Regnskabs afd."
	case "993"
	showAtt = "Administrationen"
	case else
	
	strSQL3 = "SELECT id, navn, email FROM kontaktpers WHERE id="& intAtt
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
	showAtt = oRec3("navn")
	showAttEmail = oRec3("email")
	end if
	oRec3.close
	
	if len(showAtt) <> 0 then
	showAtt = showAtt
	else
	showAtt = ""
	end if
	
	end select	
	      
	      
	      if media = "xml" then  
	      
	      'if len(trim(showAttEmail)) <> "" then
	      'xmlAtt = showAttEmail
	      'else
	      xmlAtt = showAtt
	      'end if
	      
	      'call utf_format(xmlAtt)
          'xmlAtt = utf_formatTxt
	      
	      strXML = strXML &"<com:BuyerContact>"
          strXML = strXML &"<com:ID><![CDATA["&EncodeUTF8(xmlAtt)&"]]></com:ID>" 
          strXML = strXML &"</com:BuyerContact>"
          strXML = strXML &"</com:BuyerParty>"
          end if
	
	
	'*** Afsender ***
	
		strSQL = "SELECT adresse, postnr, city, land, telefon, email, "_
		&" regnr, kkundenavn, kontonr, regnr_b, regnr_c, regnr_d, regnr_e, regnr_f, kontonr_b, kontonr_c, kontonr_d, kontonr_e, kontonr_f, cvr, bank, swift, iban,"_
        &" bank_b, swift_b, iban_b, bank_c, bank_d, bank_e, bank_f, swift_c, iban_c, swift_d, swift_e, swift_f, iban_d, iban_e, iban_f, kid, kkundenr, nace, "_
		&" fax FROM kunder WHERE kid = " & afsender
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			
            
            

            select case kontonr_sel
            case 1
            yourRegnr = oRec("regnr_b")
			yourKontonr = oRec("kontonr_b")
            yourbank = oRec("bank_b")
            yourSwift = oRec("swift_b")
            yourIban = oRec("iban_b")
            case 2
            yourRegnr = oRec("regnr_c")
			yourKontonr = oRec("kontonr_c")
            yourbank = oRec("bank_c")
            yourSwift = oRec("swift_c")
            yourIban = oRec("iban_c")
            case 3
            yourRegnr = oRec("regnr_d")
			yourKontonr = oRec("kontonr_d")
            yourbank = oRec("bank_d")
            yourSwift = oRec("swift_d")
            yourIban = oRec("iban_d")
             case 4
            yourRegnr = oRec("regnr_e")
			yourKontonr = oRec("kontonr_e")
            yourbank = oRec("bank_e")
            yourSwift = oRec("swift_e")
            yourIban = oRec("iban_e")
             case 5
            yourRegnr = oRec("regnr_f")
			yourKontonr = oRec("kontonr_f")
            yourbank = oRec("bank_f")
            yourSwift = oRec("swift_f")
            yourIban = oRec("iban_f")
            case else
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
            yourbank = oRec("bank")
            yourSwift = oRec("swift")
            yourIban = oRec("iban")
			end select
            
            
            yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")

            yourNace = oRec("nace")
			
			
			afskid = oRec("kid")
			yourFax = oRec("fax")
			yourKnr = oRec("kkundenr")
			
			if media = "xml" then
			  
			  session.LCID = 1033
			  
			  
			   call htmlparseCSV(yourAdr)
              yourAdr = htmlparseCSVtxt
			  
			  'call utf_format(yourNavn)
              'yourNavn = utf_formatTxt
              
              'call utf_format(yourAdr)
              'yourAdr = utf_formatTxt
          
             
              
               'strXML = strXML &"<com:HouseNumber>0</com:HouseNumber>" 
              
              'call utf_format(yourCity)
              'yourCity = utf_formatTxt
			  
			  strXML = strXML &"<com:SellerParty>"
              strXML = strXML &"<com:ID schemeID=""CVR"">"&yourCVR&"</com:ID>" 
              strXML = strXML &"<com:PartyName>"
              strXML = strXML &"<com:Name><![CDATA["&EncodeUTF8(yourNavn)&"]]></com:Name>" 
              strXML = strXML &"</com:PartyName>"
              strXML = strXML &"<com:Address>"
              strXML = strXML &"<com:ID>Betaling</com:ID>"
              strXML = strXML &"<com:Street><![CDATA["&EncodeUTF8(yourAdr)&"]]></com:Street>" 
              strXML = strXML &"<com:CityName><![CDATA["&EncodeUTF8(yourCity)&"]]></com:CityName>" 
              strXML = strXML &"<com:PostalZone>"& yourPostnr &"</com:PostalZone>"
              strXML = strXML &"<com:Country><com:Code listID=""ISO 3166-1"">DK</com:Code></com:Country>" 
              strXML = strXML &"</com:Address>"
              strXML = strXML &"<com:PartyTaxScheme>"
              strXML = strXML &"<com:CompanyTaxID schemeID=""CVR"">"& yourCVR &"</com:CompanyTaxID>" 
              strXML = strXML &"</com:PartyTaxScheme>"
              strXML = strXML &"</com:SellerParty>"
              
                strXML = strXML &"<com:PaymentMeans>"
                strXML = strXML &"<com:TypeCodeID>null</com:TypeCodeID>"
                
                
            if month(forfaldsdato) < 10 then
            monthffFakd = "0"& month(forfaldsdato)
            else
            monthffFakd = month(forfaldsdato)
            end if
            
            if day(forfaldsdato) < 10 then
            dayffFakd = "0"& day(forfaldsdato)
            else
            dayffFakd = day(forfaldsdato)
            end if
            
            fakforfaldsdatoXML = year(forfaldsdato) &"-"& monthffFakd &"-"& dayffFakd
                
                strXML = strXML & "<com:PaymentDueDate>"& fakforfaldsdatoXML &"</com:PaymentDueDate>" 
                
                PaymentChannelCode = "KONTOOVERFØRSEL"
                'call utf_format(PaymentChannelCode)
                'PaymentChannelCode = utf_formatTxt
                
                strXML = strXML &"<com:PaymentChannelCode>"&EncodeUTF8(PaymentChannelCode)&"</com:PaymentChannelCode>"
                
                'if lto = "epi_no" then
                'strXML = strXML &"<com:PayeeFinancialAccount><com:ID>"&yourIban&"</com:ID>"
                'else
                strXML = strXML &"<com:PayeeFinancialAccount><com:ID>"&yourKontonr&"</com:ID>"
                'end if

                strXML = strXML &"<com:TypeCode>BANK</com:TypeCode><com:FiBranch>"
                strXML = strXML &"<com:ID>"&yourRegnr&"</com:ID>"
                
                strXML = strXML &"<com:FinancialInstitution>"
                
                'if lto = "epi_no" then
                'strXML = strXML &"<com:ID>"&EncodeUTF8(yourSwift)&"</com:ID>"
                'else
                strXML = strXML &"<com:ID>null</com:ID>"
                'end if

                strXML = strXML &"<com:Name>"&EncodeUTF8(yourbank)&"</com:Name>"
                
                strXML = strXML &"</com:FinancialInstitution>"

                strXML = strXML &"</com:FiBranch></com:PayeeFinancialAccount>"
                strXML = strXML &"<com:PaymentAdvice>"
                strXML = strXML &"<com:AccountToAccount>"
                strXML = strXML &"<com:PayerNote></com:PayerNote>"
                strXML = strXML &"<com:PayeeNote></com:PayeeNote>"
                strXML = strXML &"</com:AccountToAccount>"
                strXML = strXML &"</com:PaymentAdvice>"
                strXML = strXML &"</com:PaymentMeans>"
              
              end if
			
		end if
		oRec.close
		
		
		'**** XML totol sum ***'
		if media = "xml" then
		  
		  
		  session.LCID = 1033
		  
		  strXML = strXML & "<com:TaxTotal>"
          strXML = strXML & "<com:TaxTypeCode>VAT</com:TaxTypeCode>" 
          
          strXML = strXML & "<com:TaxAmounts>"
          strXML = strXML & "<com:TaxableAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(subtotalTilmoms,2), ",", "") &"</com:TaxableAmount> "
          strXML = strXML & "<com:TaxAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intMoms,2), ",", "") &"</com:TaxAmount>" 
          strXML = strXML & "</com:TaxAmounts>"
          
          strXML = strXML & "<com:CategoryTotal>"
          strXML = strXML & "<com:RateCategoryCodeID>VAT</com:RateCategoryCodeID>"
          strXML = strXML & "<com:RatePercentNumeric>"&momssats&"</com:RatePercentNumeric>" 
          strXML = strXML & "<com:TaxAmounts>"
          strXML = strXML & "<com:TaxableAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(subtotalTilmoms,2), ",", "") &"</com:TaxableAmount>" 
          strXML = strXML & "<com:TaxAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intMoms,2), ",", "") &"</com:TaxAmount>" 
          strXML = strXML & "</com:TaxAmounts>"
          strXML = strXML & "</com:CategoryTotal>"
          strXML = strXML & "</com:TaxTotal>"
          
          strXML = strXML & "<com:LegalTotals>"
          strXML = strXML & "<com:LineExtensionTotalAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intFaktureretBelob,2), ",", "") &"</com:LineExtensionTotalAmount>" 
          strXML = strXML & "<com:ToBePaidTotalAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intFaktureretBelob+intMoms, 2), ",", "") &"</com:ToBePaidTotalAmount>"
          strXML = strXML & "</com:LegalTotals>"
		
		end if
		
		
		
		'*********************************************************************************'
		'**** Fakturaen ***'
		'*********************************************************************************'
	
		
        
		'*** fak header layout *******
        select case lto
        case "synergi1", "intranet - local"
        call fakheader() 
        case else
		call fakheader() 
	    end select


		'*** Maintable layout ********
		select case lto
        case "synergi1", "intranet - local"
        call maintable 
        case else
		call maintable 
	    end select
		
		'*** Modtager boks layout ****
		'call modtager_layout()
		
		'*** Maintable 2 layout ********
		select case lto
        case "synergi1", "intranet - local"
        call maintable_2
        case else
		call maintable_2 
	    end select
		
		'*** Afsender boks layout ****										
		'call afsender_layout()
		
		'*** Maintable 3 layout ********
         select case lto
        case "synergi1", "intranet - local"
        call maintable_3 
        case else
		call maintable_3 
	    end select
		
		
												
		'*** Vedr. (upspecificering top) ***
		'*** Vise hvis der findes jobnavn (visjobnavn ikke er slået fra), jobbesk, reknr eller det er en aftale der faktureres
		if (cint(visikkejobnavn) = 0 OR len(trim(strJobBesk)) <> 0 OR rekvnr <> "0") then
		call vedr
		else
		Response.Write "<tr><td colspan=10><br><br>&nbsp;</td></tr>"
		end if
		
		
		
		
		public fasebelialt
		'**** Udspecificering ******
		showmomsfriTxt = 0
		strFakdet = ""
		'strMedarbFakdet = ""
		strFakmat = ""
		intBelobforMoms = 0
		
        pageno = 1


		dim strFakmat, matsubtotal_fakgk
		redim matsubtotal_fakgk(1) 'strFakmat(1),
		
		
		'if jobid <> 0 then
		    
		    '** hidesumaktlinier = vis kun totalbeløb **'
		    if hidesumaktlinier <> 1 then
					
					
					'***************************************'
					'*** Fakturering af job /Aktiviteter ***'
					h = 1
					antalAktlinier = 0
					fd_aktid = 0
                    fsno = 1

					strSQL = "SELECT antal, beskrivelse, aktpris, fakid, "_
					&" enhedspris, aktid, showonfak, rabat, enhedsang, fd.valuta, fd.kurs, v.valutakode, momsfri, fase "_
					&" FROM faktura_det fd "_
					&" LEFT JOIN valutaer v ON (v.id = fd.valuta) "_
					&" WHERE fakid = "& id &" AND showonfak = 1 ORDER BY fase, fak_sortorder, beskrivelse"
					
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
						

                        fd_aktid = oRec("aktid")
						fd_valuta = oRec("valuta")
						fd_kurs = oRec("kurs")
						fd_valutaKode = oRec("valutakode")
						
						'** Pris på sum akt. i den valuta der er valgt på akt. **'
						fd_Ehpris = oRec("enhedspris")
						fd_Pris = oRec("aktpris")
						
						'** Pris på sum akt. i den valuta der er valgt på faktura. **'
						if fak_kurs <> fd_kurs then
						'aktEhpris = oRec("enhedspris") * (fak_kurs/fd_kurs)
						'else
						aktEhpris = formatnumber(oRec("enhedspris") * (fd_kurs/fak_kurs), 2)
						else
						aktEhpris = formatnumber(fd_Ehpris, 2)
						end if
						
						
						fd_momsfri = oRec("momsfri")
                        fd_fase = oRec("fase")
						
						'** aktpris er altid i dent valuta der er angivet på faktura ***'
						if len(trim(oRec("aktpris"))) <> 0 then
						aktPris = formatnumber(oRec("aktpris"),2) 
						else
						aktPris = formatnumber(0, 2)
						end if
						
						'aktEhpris = oRec("enhedspris") * (100/fak_kurs)
						'aktPris = oRec("aktpris") * (100/fak_kurs)
						
						
						if intFaktype = 1 AND media <> "xml" then
						    fd_Ehpris = -(formatnumber(fd_Ehpris, 2))
						    fd_Pris = -(formatnumber(fd_Pris, 2))
						    aktPris = -(formatnumber(aktPris, 2))
						    aktEhpris = -(formatnumber(aktEhpris, 2))
						end if
						
						rabat = oRec("rabat")
						
                        
                       'if lto = "dencker" AND cint(intKundeid) = 9 then 'Grundfos altid PC
                       'enhedsang = "PC" '"Stk. pris"
                       'else


						select case oRec("enhedsang")
                        case -1
	                    enhedsang = "" 'Ingen (skjul)
						case 0
						enhedsang = erp_txt_035 '"pr. time"
						case 1
						enhedsang = erp_txt_036 '"pr. stk."
						case 2
						enhedsang = erp_txt_037 '"pr. enhed"
						case 3
						enhedsang = erp_txt_054 '"pr. km."
						end select
						
                       'end if

					   call jquery_repl_spec(enhedsang)
                       enhedsang = jquerystrTxt
					
				    
				     
				    if cint(h) = cint(sideskiftlinier) AND cint(sideskiftlinier) <> 0 AND cint(sideskiftlinier) < 31 then
				    
				   
				   pageno = pageno + 1
				    
				    strFakdet = strFakdet &"<tr>"_
				    &"<td valign=top style='padding-right:5px; height:"& sidetoluft &"px;' colspan=8>"_
				    &"<img src=""../ill/blank.gif"" height='"& sidetoluft &"' width=0 border=0 style=""page-break-before: always;""><i>..."& txt_001 &" "& txt_059 &" "& pageno &"</i></td></tr>"
				   
				    h = 0
				    
				    end if
				   
				    
				    
				    '*** Momsfri ***'
					if cint(fd_momsfri) = 1 then
					fd_momsfriVal = "*"
					showmomsfriTxt = 1
					else
					fd_momsfriVal = ""
					end if
				    
                    tAntalTxt = ""

				    if len(trim(oRec("antal"))) <> 0 then
                        select case lto
                        case "dencker"
                        
                             'if cint(intKundeid) = 9 then 'Grundfos altid PC        
                             '   tAntal = 1
                             '   tAntalTxt = " ("& formatnumber(oRec("antal"), 2)  &")"
                             'else
                                tAntal = formatnumber(oRec("antal"), 2)
                             'end if

                        case else

                            tAntal = formatnumber(oRec("antal"), 2)
                            
                        end select
                       
				        
			
            	    else
				    tAntal = formatnumber(0, 2)
				    end if
				    
				    if len(trim(oRec("beskrivelse"))) <> 0 then
                    select case lto
                    case "dencker"
                         
                         'if cint(intKundeid) = 9 then 'Grundfos altid PC  
                         'aktBesk = oRec("beskrivelse") & tAntalTxt
                         'else
                         aktBesk = oRec("beskrivelse") & tAntalTxt
                         'end if

                    case else
                    
                    aktBesk = oRec("beskrivelse")        

                    end select

				    
				    else
				    aktBesk = ""
				    end if

                    if cint(hidefasesum) <> 1 then
                            

                            '** Gl. fase sum **'
                            if lcase(lastfdfase) <> lcase(fd_fase) AND len(trim(lastfdfase)) <> 0 AND faseshowtot = 1 then

                             strFakdet = strFakdet &"<tr>"
                             
                             
                                  if cint(hideantenh) <> 1 then
                                     select case lto 
                                     case "dencker"
                                         if cint(fak_fomr) = 16 then'Grundfos
                                         strFakdet = strFakdet &"<td colspan=3>&nbsp;</td>"
                                         else
                                         strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                         end if
                                     case else
                                     strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                     end select
                                 end if
                            
				             '<b>"& trim(replace(lastfdfase, "_", "")) &" i alt:</b>
                             strFakdet = strFakdet &"<td valign=top colspan=5 style='padding:1px 5px 10px 0px;'>&nbsp;</td>"_
                             &"<td align=right style='padding:1px 0px 10px 5px;'><b><u>"& formatnumber(fasebelialt, 2)&" "& valutaISO &"</u></b></td></tr>"


                            end if

                            '** Ny fase **'
                             if lcase(lastfdfase) <> lcase(fd_fase) AND len(trim(fd_fase)) <> 0 then
                             strFakdet = strFakdet &"<tr>"

                             if cint(hideantenh) <> 1 then
                                 select case lto 
                                 case "dencker"
                                     if cint(fak_fomr) = 16 then'Grundfos
                                     strFakdet = strFakdet &"<td align=right style='padding-right:12px;'>&nbsp;"& fsno &"</td>"
                                     strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                     else
                                     strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                     end if
                                 case else
                                 strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                 end select
                             end if

				             strFakdet = strFakdet &"<td valign=top colspan=6 style='padding:5px 5px 0px 0px;'><b>"& trim(replace(fd_fase, "_", "")) &"</b></td>"

                             if cint(fastpris) = 2 then 'commi
                                strFakdet = strFakdet &"<td valign='bottom' align=right>&nbsp;</td>"
                                strFakdet = strFakdet &"<td valign='bottom' align=right>&nbsp;</td>"
                             end if

                              strFakdet = strFakdet &"</tr>"
                            
                             fsno = fsno + 1
                             faseshowtot = 1
                             fasebelialt = 0
                             end if

                    end if
				    
				    call jq_format(aktBesk)
				    aktBesk = jq_formatTxt 
				    
				    strFakdet = strFakdet &"<tr>"

                    '** Skjul antal og enh kolonner **'
                    if cint(hideantenh) <> 1 then

                                     select case lto 
                                     case "dencker"
                                         if cint(fak_fomr) = 16 then'Grundfos
                                         strFakdet = strFakdet &"<td>&nbsp;</td>"
                                         else
                                         strFakdet = strFakdet
                                         end if
                                     case else
                                     strFakdet = strFakdet 
                                     end select


                    strFakdet = strFakdet &"<td valign=top align=right style='padding-right:10px; width:27px;'>"& tAntal &"</td>"_
					&"<td valign=top style='padding-right:5px; width:40px;'>"& enhedsang &"</td>"
                    end if

					strFakdet = strFakdet &"<td colspan=2 valign=top style='padding-left:0px;'>"& aktBesk &"</td>"_
					&"<td valign=bottom align=right style='padding-right:2px;'>"
					
					
					
					
					'*** Enh. pris og Valuta på sum akt. linie ***'
					if cint(fak_valuta) <> cint(fd_valuta) then
					strFakdet = strFakdet & formatnumber(fd_Ehpris, 2) &" "& fd_valutaKode & " ~"
				    else
					strFakdet = strFakdet & "&nbsp;"
					end if
					
					strFakdet = strFakdet &"</td>"
					
                    '** Skjul antal og enh kolonner **'
                    if cint(hideantenh) <> 1 then
                    '*** Enh. pris og Valuta på fak. ***'
					strFakdet = strFakdet &"<td valign=bottom align=right style='padding-right:2px; white-space:nowrap;'>"& formatnumber(aktEhpris, 2) &" "& valutaISO &"</td>"
					
                    else
                    strFakdet = strFakdet &"<td>&nbsp;</td>"
                    end if
					
					
					'*** Rabat ***'
					if cint(visrabatkol) = 1 then
					strFakdet = strFakdet &"<td valign=bottom align=right style='padding-right:10;'>"
					        
					        if cdbl(rabat) <> 0 then
					        rbtthis = (rabat * 100) 
					        strFakdet = strFakdet & rbtthis &" %</td>"
					        else
					        strFakdet = strFakdet &"&nbsp;</td>"
					        end if
					else
					strFakdet = strFakdet &"<td>&nbsp;</td>"
					        
					end if
					
					
					
					'strFakdet = strFakdet &"<td valign=top align=right style='padding-right:5;'>"& fd_momsfriVal &"</td>"
					
					
					
					'*** Total beløb pr. linie ***'
					strFakdet = strFakdet &"<td valign=bottom align=right>"& fd_momsfriVal &" "& formatnumber(aktPris,2) &" "& valutaISO &"</td>"


                            if cint(fastpris) = 2 then 'commi
                            strFakdet = strFakdet &"<td valign='bottom' align=right>&nbsp;</td>"
                            strFakdet = strFakdet &"<td valign='bottom' align=right>&nbsp;</td>"
                            end if

					strFakdet = strFakdet &"</tr>"
					
					'aktsubtotal = aktsubtotal + (aktPris)
					
					'Response.Write "aktsubtotal:" & aktsubtotal & "<br>"
					
                    lastfdfase = fd_fase
                    fasebelialt = fasebelialt + aktPris 

						
			        rbtthis = 0
					
					
					  if media = "xml" then
					  
					          session.LCID = 1033
					  
					          'call utf_format(aktBesk)
                              'aktBesk = utf_formatTxt
					  
					          'call utf_format(enhedsang)
                              'enhedsang = utf_formatTxt

                               IF enhedsang = "Pc." then
                               enhedsangXML = "EA"
                               else
                                enhedsangXML = enhedsang
			                   end if		  

					          strXML = strXML & "<com:InvoiceLine>"
                              strXML = strXML & "<com:ID>"&h&"</com:ID>" 
                              strXML = strXML & "<com:InvoicedQuantity unitCode='"&EncodeUTF8(enhedsangXML)&"' unitCodeListAgencyID=""n/a"">"& replace(tAntal, ",", "")&"</com:InvoicedQuantity>" 
                              strXML = strXML & "<com:LineExtensionAmount currencyID='"&valutaISO&"'>"& replace(aktPris, ",", "") &"</com:LineExtensionAmount>" 
                              strXML = strXML & "<com:Item>"
                              strXML = strXML & "<com:ID>0</com:ID>" 
                              strXML = strXML & "<com:Description><![CDATA["&EncodeUTF8(oRec("beskrivelse"))&"]]></com:Description>" 
                              strXML = strXML & "</com:Item>"
                              strXML = strXML & "<com:BasePrice>"
                              strXML = strXML & "<com:PriceAmount currencyID='"&valutaISO&"'>"& replace(aktEhpris, ",", "") &"</com:PriceAmount>" 
                              strXML = strXML & "</com:BasePrice>"
                              strXML = strXML & "</com:InvoiceLine>"

                      end if

                      h = h + 1	

						
								
								ehpris = replace(oRec("enhedspris"), ",", ".")
								
								if media <> "xml" then
								
                                if jobid <> 0 then 'kun jobfakturaer kan indeholde medarb. linier

								            '**** Medarbejder timer der vises på fakturalayout / media ***'
								            strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, "_
								            &" showonfak, medrabat, enhedsang, fms.valuta, fms.kurs, valutaKode "_
								            &" FROM fak_med_spec fms "_
								            &" LEFT JOIN valutaer v ON (v.id = fms.valuta) "_
								            &" WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND "_
								            &" enhedspris = "& ehpris &" AND fms.valuta = "& fd_valuta &" AND showonfak = 1"
								            'Response.write strSQL2
								            'Response.flush
								            oRec2.open strSQL2, oConn, 3 
								            while not oRec2.EOF 
								
								            fms_valuta = oRec("valuta")
						                    fms_kurs = oRec("kurs")
						                    fms_valutaKode = oRec("valutakode")
								
								
								            fms_Ehpris = oRec2("enhedspris")
								            fms_Belob = oRec2("beloeb")
								
								            if fak_kurs <> fd_kurs then
						                    'medEhpris = fms_Ehpris * (fak_kurs/fms_kurs)
						                    'else
						                    medEhpris = fms_Ehpris * (fms_kurs/fak_kurs)
						                    else
						                    medEhpris = fms_Ehpris
						                    end if
								
								
								
								            '** medarbejder beløb er omregnet til den valgte valuta **'
								            '** på faktura inden den lægges i DB ***'
								
								             medBelob = oRec2("beloeb")
								            'medEhpris = oRec2("enhedspris") * (100/fak_kurs)
								            'medBelob = oRec2("beloeb") * (100/fak_kurs)
								
								
								            if intFaktype = 1 AND media <> "xml" then
								            fms_Ehpris = -(formatnumber(fms_Ehpris, 2))
								            fms_Belob = -(formatnumber(fms_Belob, 2))
								            medBelob = -(formatnumber(medBelob, 2))
								            medEhpris = -(formatnumber(medEhpris, 2))
								            end if
								
								            medrabat = oRec2("medrabat")
								
								            select case oRec("enhedsang")
                                            case -1
	                                        enhedsang = "" 'Ingen (skjul)
						                    case 0
						                    enhedsang = txt_035 '"pr. time"
						                    case 1
						                    enhedsang = txt_036 '"pr. stk."
						                    case 2
						                    enhedsang = txt_037 '"pr. enhed"
						                    case 3
						                    enhedsang = txt_054 '"pr. km."
						                    end select
        								
        					                        '*** ÆØÅ **'
						                            call jquery_repl_spec(enhedsang)
                                                    enhedsang = jquerystrTxt
                                        
								
								
								                select case lto
				                                case "dencker", "tooltest", "outz", "intranet - local", "optionone"
                				    
				                                if h = 16 then '11
                				    
				                                strFakdet = strFakdet &"<tr>"_
				                                &"<td valign=top width=40 align=right style='padding-right:5px; height:100px;' colspan=8>"_
				                                &"<img src=""../ill/blank.gif"" height=100 width=0 border=0 style=""page-break-before: always;""></td></tr>"
                				   
				                                h = 0
                				    
				                                end if
				                                case else
                				    
				                                end select
								   
								   
								            strFakdet = strFakdet &"<tr>"


                                            '** Skjul antal og enh kolonner **'
                                            if cint(hideantenh) <> 1 then
								            strFakdet = strFakdet &"<td valign=top align=right style='padding-right:10px; width:27px;'><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								            &"<td valign=top width=40 style='padding-right:5px;'><font class=megetlillesort>"& enhedsang &"</td>"
                                            end if

								            strFakdet = strFakdet &"<td colspan=2 valign=top style='padding-left:0px;'><font class=megetlillesort>" & oRec2("tekst") &"</td>"_
								            &"<td valign=top align=right style='padding-right:2px;'><font class=megetlillesort>"
					
					                        '*** Enh. pris og Valuta på medarb. linie ***'
					                        if cint(fak_valuta) <> cint(fms_valuta) then
					                        strFakdet = strFakdet & formatnumber(fms_Ehpris, 2) &" "& fms_valutaKode & " ~"
					                        else
					                        strFakdet = strFakdet & "&nbsp;"
					                        end if
            					
					                        strFakdet = strFakdet &"</td>"
								
                                             '** Skjul antal og enh kolonner **'
                                            if cint(hideantenh) <> 1 then
								            '*** Enh. pris og Valuta på medarb. fra Faktura valuta ***'
								            strFakdet = strFakdet &"<td valign=top align=right style='padding-right:2px;'><font class=megetlillesort>"& formatnumber(medEhpris, 2) &" "& valutaISO &"</td>"
                                            else
                                            strFakdet = strFakdet &"<td>&nbsp;</td>"
                                            end if
								
								
								            '*** Rabat ***'
								            if cint(visrabatkol) = 1 then
								            strFakdet = strFakdet &"<td valign=top align=right style='padding-right:10px;'>"
								
								                if cdbl(medrabat) <> 0 AND cint(visrabatkol) = 1 then
								                strFakdet = strFakdet &"<font class=megetlillesort>" & (medrabat * 100) &" %</td>"
								                else
								                strFakdet = strFakdet &"&nbsp;</td>"
								                end if
								            else
								            strFakdet = strFakdet &"<td>&nbsp;</td>"
								
								            end if
								
								            '*** Total beløb pr. linie ***'
								            strFakdet = strFakdet &"<td valign='top' align=right><font class=megetlillesort>"& formatnumber(medBelob, 2) &" "& valutaISO &"</td>"_
								            &"</tr>"
								
								            h = h + 1
								
								            oRec2.movenext
								            wend
								            oRec2.close 
								
                                end if 'jobid

								end if '** media XML

                

                                '*** Hvis materialer vises efter hver aktivitet *****'
                                if jobid <> 0 then

                            

                                    	if cint(showmatasgrp) = 2 then
                                            
                                            call matlinjer
                                            strFakdet = strFakdet & strFakmat


                                            
                                        end if

                                end if


								
					antalAktlinier = antalAktlinier + 1				
					lastaktid = oRec("aktid")			
					oRec.movenext
					wend
					oRec.close 
					

                    if jobid <> 0 then
					
					    '** Gl. fase sum **'
                        if cint(hidefasesum) <> 1 then

                            if h > 0 AND len(trim(fd_fase)) <> 0 AND faseshowtot = 1 then

                                strFakdet = strFakdet &"<tr>"

                                 if cint(hideantenh) <> 1 then
                                     select case lto 
                                     case "dencker"
                                         if cint(fak_fomr) = 16 then'Grundfos
                                         strFakdet = strFakdet &"<td colspan=3>&nbsp;</td>"
                                         else
                                         strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                         end if
                                     case else
                                     strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                     end select
                                 end if


                                '"& trim(replace(lastfdfase, "_", "")) &" i alt:
				                strFakdet = strFakdet &"<td valign=top colspan=5 style='padding:1px 5px 10px 0px; color:#999999;'>&nbsp;</td>"_
                                &"<td align=right style='padding:1px 0px 10px 5px;'><b><u>"&formatnumber(fasebelialt, 2)&" "& valutaISO &"</u></b></td>"


                               if cint(fastpris) = 2 then 'commi
                                   strFakdet = strFakdet &"<td valign='bottom' align=right>&nbsp;</td>"
                                   strFakdet = strFakdet &"<td valign='bottom' align=right>&nbsp;</td>"
                               end if


                                 strFakdet = strFakdet &"</tr>"

                            end if

                        end if

                    end if
					
					
					'if len(medarbejdertimer) <> 0 then
					'medarbejdertimer = medarbejdertimer
					'else
					'medarbejdertimer = 0
					'end if
					
					'if len(medarbejdersum) <> 0 then
					'medarbejdersum = medarbejdersum
					'else
					'medarbejdersum = 0
					'end if
					
					if medarbejdertimer <> intFaktureretTimer then
					bgthis = "#ffffe1"
					else
					bgthis = "#ffffff"
					end if
					
					'if aktsubtotal <> 0 then
					'aktsubtotal = aktsubtotal 
					'else
					'aktsubtotal = 0
					'end if
					
					        
					'intBelobforMoms = aktsubtotal 
					 
					 
					           
					           
					        '********************************************'
					        '*** Materialer *****'
					        '*** m/u moms m = 0 med moms m = 1 u moms ***'
					        '********************************************'
                            fd_aktid = 0
                            
					        call matlinjer
					        
					        
					        
					    
		          end if 'hidesumaktlinier			
		
		'else '**** Fakturering af Aftaler ***
		
				
				
				
		
		
		

        '*** Skriver faktura linjer ****'
		'**** Aktiviteter / timer ***'
        %>
        <table cellspacing="0" cellpadding="1" border="0" width="<%=gblWdt%>">
        <%
		if len(trim(strFakdet)) <> 0 then
 		
 		    aktmat = 1
 		    call udspecificering(strFakdet, aktmat, antalAktlinier, fastpris)
		
		
		    if jobid <> 0 then
		    '**** Sub-tot. aktiviteter ***'
		        select case lto
		        case "essens"
		        case else
		            'call subtotakt()
		        end select
		    end if
		
		
		end if
		    
		    
    		
		    '*** Materialer ***'
                if len(trim(strFakmat)) <> 0 then
		        
		        
    		        '**** Luft før materialer ***'
    		        select case lto
		            case "essens"
		            case else
		            'Response.Write "<tr><td colspan=8 height=20>&nbsp;</td></tr>"
		            end select
		    
		   
		                aktmat = 2 
		                call udspecificering(strFakmat, aktmat, antalAktlinier, fastpris)
    		    
    		    
    		        '**** Sub-tot. materialer ***'
    		        select case lto
		            case "essens"
		            case else
		            'call subtotmat(0)
		            end select
		    
		        end if
            
            '** afslutter alle udspecificinger master TABLE (startes før første udspec)
            %>
            </table>
            <%
		    
		    
		    '*** Rykker gebyrer ***'
		    call rykkergebyrer()
		    
		    '**** Totaler beløb til moms ***'
		    call totogmoms()
		    
		



		    '*** Materialer u/moms ***'
            '*** Bruges ikke mere ***'
            qltf = 0
            if qltf = 1000 then
		    if len(trim(strFakmat)) <> 0 then
		    'Response.Write "<tr><td colspan=8 height=20>&nbsp;</td></tr>"
            aktmat = 3
		    call udspecificering(strFakmat, aktmat, antalAktlinier, fastpris) 'strFakmat(1)
    		'**** Sub-tot. materialer ***'
    		    
    		    select case lto
		        case "essens"
		        case else
		        call subtotmat(1)
		        end select
		    
		    end if
            end if
		
		    
		
		
		
		'**** Totaler ***'
		call totalbelob_fakgk()
		
		
		'*** Komm. og bet. betingelser ***'
		call komogbetbetingelser
		
		
		Response.Write "</div>" 'fakturaside
		
     
		
		
		stDatoKri = request("FM_start_aar_ival") &"/"& request("FM_start_mrd_ival") &"/"& request("FM_start_dag_ival")  
		slutdato = request("FM_slut_aar_ival") &"/"& request("FM_slut_mrd_ival") &"/"& request("FM_slut_dag_ival")  
		
		
		'*** Joblog ****
		if visjoblog = 1 then%>
		
		<!-- joblog -->
		
	    <div id="joblog" style="position:relative; width:<%=gblWdt%>px; left:10px; top:250px; visibility:visible; page-break-before:always; background-color:#FFFFFF; border:0px #8caae6 solid; padding:10px 10px 10px 10px;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td colspan=6>
                <%'if lto = "acc" then
            'Response.write "her " & jobid &", "& stdatoKri &","& slutdato &","& aftid 
            'end if %>
                <h4><%=txt_038 %></h4></td>
		</tr>
		
		
		<%  
           
		call joblog(jobid, stdatoKri, slutdato, aftid)
		%>
		
		</table>
		</div>
		
		<br><br>&nbsp;
		<!-- Joblog slut -->
		<%end if 	
		
		
		
		
		'*** Matlog ****
		if vismatlog = 1 then%>
		
		<!-- matlog -->
		
		<!--<div id="matlog" style="position:relative; width:640px; left:10px; top:50px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">-->
        <div id="matlog" style="position:relative; width:<%=gblWdt%>px; left:10px; top:250px; visibility:visible; border:0px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		<table border=0 cellspacing=0 width=620 cellpadding=0 style="page-break-before:always;">
		<tr>
			<td><h4><%=txt_039 %></h4></td>
		</tr>
		</table>
		
		<table width=100% cellspacing=0 cellpadding=0 border=0>
		<tr>
	                            <td><b><%=txt_042 %></b></td>
	                            <td><b><%=txt_043 %></b></td>
	                            <td><b><%=txt_026 %></b></td>
	                            <td align=right><b><%=txt_013 %></b></td>
	                            <td align=right><b><%=txt_020 %></b></td>
	                            <td align=right><b><%=txt_044 %></b></td>
	                            <td align=right><b><%=txt_017 %></b></td>
	                        </tr>
		
		<%
		call matlog(jobid, stdatoKri, slutdato, aftid)
		%>
		
		</table>
		</div>
		<br><br>&nbsp;
		
		
		<!-- matlog slut -->
		<%end if %>	
		
		</div>
		<br /><br />
        &nbsp;
											
		<%
		'**** Højreside (ikke printbar) ***
		'** IVal bruges til at fp joblog med ud ***'
		
		
		if media = "n" then %>
		<div id=funktioner style="position:absolute; z-index:200; top:102px; width:300px; left:1070px; border:0px #8caae6 DASHED; background-color:pink; padding:10px 10px 10px 10px;">
		<b><%=erp_txt_222 %></b> <br />
        
        <a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&media=j&id=<%=id%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu><%=erp_txt_223 %></a>
		<!--
		<a href="erp_fak_godkendt_2007.asp?media=j&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&kid=<%=kid%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu>Print</a>
        -->  
        &nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="erp_make_pdf.asp?media=pdf&lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu><%=erp_txt_224 %></a>
            <!-- &FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%> -->
		
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="#" class=vmenu id="oioxml"><%=erp_txt_225 %></a> 
        
        <%select case lto
	    case "epi", "epi_no", "epi_sta", "outz", "intranet - local", "synergi1", "epi_ab", "rek"
        showvans = 1
        %>
        (Afsend XML fil til E-posthus)
	    <%
        case else
        showvans = 0
	    end select
	    %>

		 
        <form action="#" method="post">
       <!-- <input id="Xxmlurl" type="hidden" value="erp_fak_godkendt_2007.asp?media=xml&lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>" />-->
        <input id="xmlurl" type="hidden" value="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&media=xml&lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>" />
        <input id="showvans" type="hidden" value="<%=showvans %>" />
        </form>

       
	    <br /><br />
	    <%select case lto
	    case "execon", "immenso"', "intranet - local"
	    
	    case else
	    %>
	   
	    <b><%=erp_txt_226 %></b><br />
            <form><%=erp_txt_227 %> <input type="text" id="FM_sogkpers" /><input id="bt_sogkpers" type="button" value=">>" /></form>
		
            <div id="kpers_div"><%
		
		'pdfurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\faktura_"&lto&"_"&faknr&".pdf"
	    
		
		'strSQL = "SELECT navn, email, titel FROM kontaktpers WHERE kundeid = " & kid
        strSQL = "SELECT kp.navn, kp.email, kp.titel, k.kid, k.kkundenavn FROM kontaktpers AS kp "_
        &" LEFT JOIN kunder AS k on (k.kid = kp.kundeid) "_
        &" WHERE kundeid = "& kid &" ORDER BY k.kkundenavn, kp.navn" 
		
         lastKid = 0
		'Response.Write strSQL 
		'Response.Flush
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF

        if lastKid <> oRec("kid") then
        Response.Write "<br><br><b>"& oRec("kkundenavn") & "</b><br>"
        end if

        Response.Write "<i>"& oRec("navn") & "</i> "
        
        if len(trim(oRec("titel"))) <> 0 then
        Response.Write " ("& oRec("titel") &") "
        end if
        
        
        Response.Write " <a href='mailto:"&oRec("email")&"&subject=Faktura: "& varFaknr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
        lastKid = oRec("kid") 

        oRec.movenext
        wend
        oRec.close
		%>
        </div>
		<br /><%=erp_txt_228 %>
		
		<%end select %>	
		
		<br /><br />
		<table cellspacing=2 cellpadding=1 border=0 width=290 bgcolor="#ffffff">
		
		<tr>
		<%
		 if intStatus = 0 then
	    %>
	    <td width=120 valign=top style="padding-top:6px;">
	    <a href="Javascript:popUp('godkendfaktura.asp?fakid=<%=id %>','400','250','50','50');" target="_self"  class=vmenu>
        <%=erp_txt_229 %></a>
	   </td>
	    <td valign=top width=170>
	     <a href="Javascript:popUp('godkendfaktura.asp?fakid=<%=id %>','400','250','50','50');" target="_self"  class=vmenu>
              <img src="../ill/godkend_pil.gif" border="0" />  </a>  </td>
          </tr>
          <tr><td colspan=2 valign=top class=lille>
        
        <%select case lto
	    case "execon", "immenso", "intranet - local" 
	    Response.Write "Afsender mail med PDF af faktura til bogholderiet."
	    case else
	    Response.Write "&nbsp;"
	    end select%>
              
              
	    </td>
	    
	    <% 
	    else
	    %>
	    
	
	   <td><img src="../ill/godkend_icon.gif" border="0" /></td>
	   <td><font color=forestgreen><b><%=erp_txt_095 %></b></font></td>
	      </tr>
		
	     
	    <%
	    end if
		%>
		</tr>
		</table>
		
		</div>
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;Excel.-->
	    <%end if
	    
	    
	    
	    
	   
	    
	    
	    
	    if media = "j" then
	    Response.Write("<script language=""JavaScript"">window.print();</script>")
	    end if
	    
		if media = "xml" then
		
		
    	session.LCID = 1033  
    	strXML = strXML & "</Invoice>"
	    
	    ext = "xml"
		
         call TimeOutVersion()

          if cdbl(id) <> 0 then 
         
             strknavnSQL = "SELECT kkundenavn, fid FROM fakturaer LEFT JOIN kunder ON (kid = fakadr) WHERE fid = " & id
             oRec.open strknavnSQL, oConn, 3
             
              strknavn = "xx"
             if not oRec.EOF then
             strknavn = oRec("kkundenavn") 
             
             call kundenavnPDF(strknavn)
             strknavn = strKundenavnPDFtxt

             end if
             oRec.close

             

         else
         strknavn = "oo"
         end if

	
	    Set objFSO = server.createobject("Scripting.FileSystemObject")
    	
    	file = "faktura_xml_"&lto&"_"&varFaknr&"_"&strknavn&"."& ext
    	
	    if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\"&toVer&"\timereg\erp_fak_godkendt_2007.asp" then
    							
		    Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", True, False)
		    Set objNewFile = nothing
		    Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", 8,-1)
    	
	    else
    		
		    Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", True, false)
		    Set objNewFile = nothing
		    Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", 8,-1)


            if cint(vans) = 1 then
                '***  VANS CSC ***
                select case lto
	            case "outz", "epi", "synergi1", "epi_no", "epi_sta", "epi_ab" , "rek"

                     if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                        
                              'Set objNewFileVans = objFSO.createTextFile("d:\webserver\CSC\CSC VANS\TIL-VANS\"& file &"", True, false)
                                Set objNewFileVans = objFSO.createTextFile("C:\Program Files\CSC\CSC VANS Transfer Service\Outbound\Invoice\INPUT\"& file &"", True, false)
		                        Set objNewFileVans = nothing
		                        'Set objFVans = objFSO.OpenTextFile("d:\webserver\CSC\CSC VANS\TIL-VANS\"& file &"", 8,-1)
                                Set objFVans = objFSO.OpenTextFile("C:\Program Files\CSC\CSC VANS Transfer Service\Outbound\Invoice\INPUT\"& file &"", 8,-1)

                     else

                                 'Set objNewFileVans = objFSO.createTextFile("d:\webserver\CSC\CSC VANS\TIL-VANS\"& file &"", True, false)
                                Set objNewFileVans = objFSO.createTextFile("C:\Program Files\CSC\CSC VANS Transfer Service - 64 Bit\Outbound\Invoice\INPUT\"& file &"", True, false)
		                        Set objNewFileVans = nothing
		                        'Set objFVans = objFSO.OpenTextFile("d:\webserver\CSC\CSC VANS\TIL-VANS\"& file &"", 8,-1)
                                Set objFVans = objFSO.OpenTextFile("C:\Program Files\CSC\CSC VANS Transfer Service - 64 Bit\Outbound\Invoice\INPUT\"& file &"", 8,-1)

                     end if

                   
                    

                    objFVans.WriteLine(strXML)
	                objFVans.close

                end select
            end if

    		
	    end if
    	
    	
		
        objF.WriteLine(strXML)
	    objF.close

	    Response.redirect "../inc/upload/"&lto&"/"& file &""
	    
	    end if
		
		
		        		        
%>


