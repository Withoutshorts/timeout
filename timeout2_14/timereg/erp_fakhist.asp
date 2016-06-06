<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%

call TimeOutVersion()

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	menu = "erp"
	
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	func = request("func")
    thisfile = "erp_tilfakturering.asp"
    print = request("print")
	
	select case func 
	case "opdbetalt"
	
	
	'Response.Write request("erfakbetalt") & "<br>"
	
	fids = split(trim(request("erfakbetalt")), ", ")
	thisKom = split(trim(request("FM_fakbetkom")), ", #, ")


    'Response.write request("overfortERP") & "<br>"

    overfortERP = split(trim(request("overfortERP")), ";")
    lukjob = split(trim(request("lukjob")), ";")

	for x = 0 to UBOUND(fids)
	
	thisval = trim(request("fakid_"&fids(x)&""))


    overfortERP(x) = replace(overfortERP(x), ",", "")
    overfortERP(x) = replace(overfortERP(x), " ", "")
    
    if overfortERP(x) = "1" then
    overfortERPval = 1
    else
    overfortERPval = 0
    end if

    lukjob(x) = replace(lukjob(x), ",", "")
    lukjob(x) = replace(lukjob(x), " ", "")
	

    
    'Response.write "<br>"& x &"=="& overfortERP(x) & "<br>"


	'*** Opdaterer fakturaer *****'
	strSQL = "UPDATE fakturaer SET erfakbetalt = "& thisval &", fakbetkom = '"& replace(thisKom(x), "'", "''") &"', overfort_erp = "& overfortERPval &" WHERE fid = "& fids(x)
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)

    '** Lukker job ****'
    if lukjob(x) <> "0" ANd len(trim(lukjob(x))) <> 0 then 

    '*** Sæt lukkedato (skal være før det skifter status) ***'
    call lukkedato(lukjob(x), 0)

    strSQL = "UPDATE job SET jobstatus = 0 WHERE id = "& lukjob(x) 
	'Response.Write strSQL & "<br>"
	
    oConn.execute(strSQL)

                
                  '**** SyncDatoer ****'
                jobnrThis = 0
                jobStatus = 0
                syncslutdato = 0
                strSQLj = "SELECT jobnr, syncslutdato FROM job WHERE id =  "& lukjob(x) 
                oRec5.open strSQLj, oConn, 3 
                if not oRec5.EOF then

                jobnrThis = oRec5("jobnr")
                syncslutdato = oRec5("syncslutdato")

                end if
                oRec5.close

                if jobStatus = 0 AND syncslutdato = 1 then
                    call syncJobSlutDato(lukjob(x), jobnrThis, syncslutdato)
                end if



    end if
	
	next
	
	'Response.end
	
	Response.redirect "erp_fakhist.asp?FM_job="&request("jobid")&"&FM_aftale="&request("aftid")&"&FM_sog="&request("FM_sog")&"&sort="&request("sort")
   
	
	
	case else
	


    if len(trim(request("sogFaknrOrJob"))) <> 0 then
    sogFaknrOrJob = request("sogFaknrOrJob")
            
            if sogFaknrOrJob = 1 then
            sogFaknrOrJobCHK1 = "CHECKED"
            else
            sogFaknrOrJobCHK0 = "CHECKED"
            end if
    else
            
            if request.cookies("erp")("sogFaknrOrJob")  <> "" then
                
                
                sogFaknrOrJob = request.cookies("erp")("sogFaknrOrJob")    

                if sogFaknrOrJob = 1 then
                sogFaknrOrJobCHK1 = "CHECKED"
                else
                sogFaknrOrJobCHK0 = "CHECKED"
                end if

            else
            sogFaknrOrJob = 0
            sogFaknrOrJobCHK0 = "CHECKED"
            end if
     end if
	
    response.cookies("erp")("sogFaknrOrJob") = sogFaknrOrJob


	
	if print <> "j" then
	
	dTop = 102
	dLeft = 90 
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    


	<script type="text/javascript">

	    $(document).ready(function() {

         

	        $("#FM_kunde").change(function() {

	        $("#FM_sog").val('')
	        //$("#fakhist").submit()

	        });

	        $("#FM_aftale").change(function() {

	        $("#FM_sog").val('')
	        $("#fakhist").submit()

	        });

	        $("#FM_job").change(function() {

	        $("#FM_sog").val('')
	        $("#fakhist").submit()

	        });


         

             $("#alleerp").click(function() {

           
            
             if ($("#alleerp").is(':checked') == true) {
                $(".overfortERP").attr('checked', true);
            } else {
                $(".overfortERP").attr('checked', false);

            }

        });  
             
              $("#alleluk").click(function() {

           
            
             if ($("#alleluk").is(':checked') == true) {
                $(".lukjob").attr('checked', true);
            } else {
                $(".lukjob").attr('checked', false);

            }

              });




              $("#FM_viskun_fak").click(function () {

                      if ($("#FM_viskun_fak").is(':checked') == true) {
                      $("#FM_viskun_kre").attr('checked', false);

                 }

              });



              $("#FM_viskun_kre").click(function () {

                  if ($("#FM_viskun_kre").is(':checked') == true) {
                      $("#FM_viskun_fak").attr('checked', false);

                  }

              });

              $("#FM_viskun_godkendte").click(function () {

                  if ($("#FM_viskun_godkendte").is(':checked') == true) {
                      $("#FM_viskun_kladder").attr('checked', false);

                  }

              });


              $("#FM_viskun_kladder").click(function () {

                  if ($("#FM_viskun_kladder").is(':checked') == true) {
                      $("#FM_viskun_godkendte").attr('checked', false);

                  }

              });
          
          


	    });
	
	</script>
	
	
	<!--include file="../inc/regular/topmenu_inc.asp"-->
	
    <!--
    <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call erpmainmenu(1)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call erptopmenu()
	%>
	</div>

    -->
	
	<%
    
        call menu_2014()    
        
    else 
	
	dTop = 20
	dLeft = 20 
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<%end if %>
	
	
	<div id="sindhold" style="position:absolute; left:<%=dLeft %>px; top:<%=dTop %>px; visibility:visible;">
	<%
	'**********************************************************
	'**************** Job til fakturering *********************
	'**********************************************************
	
	
	
	
	
	
	''ptop = 0
	'pleft = 0
	'pwdt = 852
	
    oskrift = erp_txt_230

	'call filterheader(ptop,pleft,pwdt,pTxt) 
     call filterheader_2013(0,0,840,oskrift)%>
	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<form action="erp_fakhist.asp" id="fakhist" method="POST">
	<tr>
	<td valign=top colspan="2">
	
    <% 
    if len(request("FM_kunde")) <> 0 then
			
			'if request("FM_kunde") = 0 then
			'kundeid = 0
			'sqlKundeKri = "t.tknr <> 0"
			'sqlKundeKri2 = "k.kid <> 0"
			'else
			kundeid = request("FM_kunde")
			'sqlKundeKri = "t.tknr = "& valgtKunde &""
			'sqlKundeKri2 = "k.kid = "& valgtKunde &""
			'end if
			
	else
		
			if len(request.cookies("erp")("kid")) <> 0 AND request.cookies("erp")("kid") <> 0 then
			kundeid = request.cookies("erp")("kid")
			'sqlKundeKri = "t.tknr = "& valgtKunde &""
			'sqlKundeKri2 = "k.kid = "& valgtKunde &""
			else
			kundeid = 0
			'sqlKundeKri = "t.tknr <> 0"
			'sqlKundeKri2 = "k.kid <> 0"
			end if
			
	end if
	
	
	response.Cookies("erp")("kid") = kundeid
	response.Cookies("erp").expires = date + 60
	
	if len(request("FM_sog")) <> 0 AND request("FM_sog") <> "9999" then
	showSog = request("FM_sog")
	else
	showSog = ""
	end if
	
	if len(request("FM_job")) <> 0 then
	jobid = request("FM_job")
	else
	jobid = 0
	end if
	
	if len(request("FM_aftale")) <> 0 then
	aftid = request("FM_aftale")
	else
	aftid = 0
	end if
	
	
	
	if len(trim(showSog)) <> 0 then
	aftid = 0
	jobid = 0
	kundeid = 0
	end if
	
	select case jobid
	case -1
	vlgtJob = "Ingen"
	case 0
	vlgtJob = "Alle"
	end select 
	
	select case aftid
	case -1
	vlgtAft = "Ingen"
	case 0
	vlgtAft = "Alle"
	end select 
	
	if len(trim(request("FM_jobikkedelafaftale"))) <> 0 then
	jidaftCHK = "CHECKED"
	jidaft = 1
	else
	jidaftCHK = ""
	jidaft = 0
	end if

    
        
        
   

    '** Kun ikke betalte
    if (len(request("FM_viskunabne")) <> 0 AND request("FM_viskunabne") <> "0") then 'er der søgt via form then
	viskunabne = 1
    response.cookies("erp")("kunabne") = "1"
	else

        if len(trim(request("FM_kunde"))) <> 0 then

        viskunabne = 0
        response.cookies("erp")("kunabne") = "0"

        else

        if request.cookies("erp")("kunabne") = "1" then
        viskunabne = 1
        response.cookies("erp")("kunabne") = "1"
        else 
	    viskunabne = 0
        response.cookies("erp")("kunabne") = "0"
	    end if

        end if

    end if


        if cint(viskunabne) <> 0 then
        viskunabneCHK = "CHECKED"
        else
        viskunabneCHK = ""
        end if
	

    
    '*** kun interne
    if (len(trim(request("FM_viskun_interne"))) <> 0 AND request("FM_viskun_interne") <> "0") then
	
    viskintCHK = "CHECKED"
	viskint = 1
    response.cookies("erp")("kunint") = "1"
	
   else
	

        if len(trim(request("FM_kunde"))) <> 0 then

        viskintCHK = ""
	    viskint = 0
        response.cookies("erp")("kunint") = "0"

        else

        if request.cookies("erp")("kunint") = "1" then
        viskintCHK = "CHECKED"
	    viskint = 1
        else
        viskintCHK = ""
	    viskint = 0
        response.cookies("erp")("kunint") = "0"
        end if

        end if

	end if



    '** Kun ikke tidligere overført til ERP
    if (len(trim(request("FM_viskun_ikkecsv"))) <> 0 AND request("FM_viskun_ikkecsv") <> "0") then
	viskicsvCHK = "CHECKED"
	viskicsv = 1
    response.cookies("erp")("kunikkecsv") = "1"

	else

            if len(trim(request("FM_kunde"))) <> 0 then

             viskicsvCHK = ""
	        viskicsv = 0
            response.cookies("erp")("kunikkecsv") = "0"

            else

                  if request.cookies("erp")("kunikkecsv") = "1" then
                    viskicsvCHK = "CHECKED"
	                viskicsv = 1
                   else
	                viskicsvCHK = ""
	                viskicsv = 0
                    response.cookies("erp")("kunikkecsv") = "0"
                end if

            end if

	end if






    '** Kun godkendte
     if (len(trim(request("FM_viskun_godkendte"))) <> 0 AND request("FM_viskun_godkendte") <> "0") then
	    viskungkCHK = "CHECKED"
	    viskungk = 1
        response.cookies("erp")("kunikkegk") = "1"
	else


        if len(trim(request("FM_kunde"))) <> 0 then

          viskungkCHK = ""
	    viskungk = 0
        response.cookies("erp")("kunikkegk") = "0"


        else


             if request.cookies("erp")("kunikkegk") = "1" then
            viskungkCHK = "CHECKED"
	        viskungk = 1
            else
	        viskungkCHK = ""
	        viskungk = 0
            response.cookies("erp")("kunikkegk") = "0"
            end if

        end if


	end if






    '*** Kun kladder
     if (len(trim(request("FM_viskun_kladder"))) <> 0 AND request("FM_viskun_kladder") <> "0") then
	    viskunklCHK = "CHECKED"
	    viskunkl = 1
        response.cookies("erp")("kunikkekl") = "1"
	else

        if len(trim(request("FM_kunde"))) <> 0 then

            viskunklCHK = ""
	        viskunkl = 0
            response.cookies("erp")("kunikkekl") = "0"


        else

            if request.cookies("erp")("kunikkekl") = "1" then
            viskunklCHK = "CHECKED"
	        viskunkl = 1
            else
	        viskunklCHK = ""
	        viskunkl = 0
            response.cookies("erp")("kunikkekl") = "0"
            end if

        end if

	end if



     '*** Kun Fak
     if (len(trim(request("FM_viskun_fak"))) <> 0 AND request("FM_viskun_fak") <> "0") then
	    viskunFakCHK = "CHECKED"
	    viskunFak = 1
        response.cookies("erp")("viskunFak") = "1"
	else

        if len(trim(request("FM_kunde"))) <> 0 then

            viskunFakCHK = ""
	        viskunFak = 0
            response.cookies("erp")("viskunFak") = "0"


        else

            if request.cookies("erp")("viskunFak") = "1" then
            viskunFakCHK = "CHECKED"
	        viskunFak = 1
            else
	        viskunFakCHK = ""
	        viskunFak = 0
            response.cookies("erp")("viskunFak") = "0"
            end if

        end if

	end if



    '** Vis kun KRE
    if (len(trim(request("FM_viskun_kre"))) <> 0 AND request("FM_viskun_kre") <> "0") then
	    viskunKreCHK = "CHECKED"
	    viskunKre = 1
        response.cookies("erp")("viskunKre") = "1"
	else

        if len(trim(request("FM_kunde"))) <> 0 then

            viskunKreCHK = ""
	        viskunKre = 0
            response.cookies("erp")("viskunKre") = "0"


        else

            if request.cookies("erp")("viskunKre") = "1" then
            viskunKreCHK = "CHECKED"
	        viskunKre = 1
            else
	        viskunKreCHK = ""
	        viskunKre = 0
            response.cookies("erp")("viskunKre") = "0"
            end if

        end if

	end if



        




   if len(trim(request("FM_ignper"))) <> 0 AND request("FM_ignper") <> "0" then
	ignperCHK = "CHECKED"
	ignper = 1
	else
	ignperCHK = ""
	ignper = 0
	end if


    

     

    

    
	
    
    %>
    <b><%=erp_txt_231%></b><br />
        <%
        
        ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
        
        strSQL = "SELECT Kkundenavn, Kkundenr, kid, COUNT(f.fid) AS antalfak, f.fakadr FROM fakturaer AS f "_
				&" LEFT JOIN kunder AS k ON (kid = f.fakadr) "_
				&" WHERE "& ketypeKri &" "& strKundeKri &" AND f.fid <> 0 AND shadowcopy = 0 AND kid IS NOT NULL GROUP BY f.fakadr ORDER BY Kkundenavn"
				
                'if lto = "essens" then
                'Response.write strSQL
                'Response.end
                'end if
        
        if print <> "j" then 
        %>
        <select id="FM_kunde" name="FM_kunde" style="width:450px;" size=12> 
        
        <%if cint(kundeid) = 0 then 
        isAllSel = "SELECTED"
        else
        isAllSel = ""
        end if%>

		<option value="0" <%=isAllSel %>>Alle</option>
		<%end if
		
		
				

				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = oRec("kid") then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("kkundenr") %>) - fak: <%=oRec("antalfak") %></option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		        </select><br>
		        <% 
		        else
		        %>
				<%=vlgtKunde %>
				<% 
				end if
				%>
	    
	    <!--
	    <b>Kontaktansvarlige (medarbejder):</b><br> <select name="FM_kans" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
	    </select>-->
	    
	    
	    
	    </td>
	    
	    <%
	    if len(trim(request("sort"))) <> 0 then
	    sort = request("sort")
	    

        sortCHK4 = ""
        sortCHK3 = "" 
	    sortCHK2 = ""
	    sortCHK1 = ""
	            
	            select case sort
	            case 2
	           
	            sortCHK2 = "CHECKED"
	            sortOrderKri = "f.fakdato DESC, f.faknr"
	            case 3
	            sortCHK3 = "CHECKED"
	             sortOrderKri = "f.b_dato DESC, f.faknr"
                case 4
	            sortCHK4 = "CHECKED"
	            sortOrderKri = "f.labeldato DESC, f.faknr"
	            case else
	            sortCHK1 = "CHECKED"
	            sortOrderKri = "f.faknr"
	            end select
	            
	   
	    else
          select case lto
          case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab"
            sort = 4
            sortCHK4 = "CHECKED"
	        sortCHK3 = ""
	        sortCHK2 = ""
	        sortCHK1 = ""
	        sortOrderKri = "f.labeldato DESC, f.faknr"
          case else
	        sort = 2
            sortCHK4 = ""
	        sortCHK3 = ""
	        sortCHK2 = "CHECKED"
	        sortCHK1 = ""
	        sortOrderKri = "f.fakdato DESC, f.faknr"
         end select
	    end if
	    
	    
	    'if request("showlabel") = "1" OR (request.Cookies("erp")("showlabel") = "1" AND len(trim(request("sort"))) = 0) then
	    'showlabel = 1
	    'showlabelCHK1 = "CHECKED"
	    'showlabelCHK0 = ""
	    'else
	    'showlabel = 0
	    'showlabelCHK1 = ""
	    'showlabelCHK0 = "CHECKED"
	    'end if
	    
	    'response.cookies("erp")("showlabel") = showlabel
        
       
            
            
        if print <> "j" then%>
        <td style="padding:2px 2px 2px 20px;" valign=top>
	    <b><%=erp_txt_154 %></b> <%=replace(erp_txt_232, "|", "&") %><br /> 
	    <!--#include file="inc/weekselector_s.asp"-->&nbsp;<br />
	    
        <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
        <br />
        <b><%=erp_txt_233 %></b><br />
        
        <input id="sort" name="sort" type="radio" value="1" <%=sortCHK1 %> /><%=erp_txt_234 %><br /> 
        <input id="sort" name="sort" type="radio" value="2" <%=sortCHK2 %> /><%=erp_txt_235 %><br />
        <input id="sort" name="sort" type="radio" value="4" <%=sortCHK4 %> /><%=erp_txt_236 %><br />
        <input id="sort" name="sort" type="radio" value="3" <%=sortCHK3 %> /><%=erp_txt_237 %><br />
        
            <!--<br />
            <br />Vis fakturaer der benytter <b>faktura "label" dato</b>: <br />
            <input id="showlabel" name="showlabel" value="1" type="radio" <=showlabelCHK1 %> /> Kun hvis <b>label-dato</b> ligger i den valgte periode.<br />
            <input id="showlabel" name="showlabel" type="radio" value="0" <=showlabelCHK0 %> /> Uanset om <b>label-dato</b> ligger i den valgte per.
            
            -->
            
            <br />
            
            <b><%=erp_txt_238 %></b><br />
            <input id="FM_viskunabne" name="FM_viskunabne" value="1" type="checkbox" <%=viskunabneCHK%> /> <%=erp_txt_239 %>
            


            <br />
            <input id="Checkbox1" name="FM_viskun_interne" value="1" type="checkbox" <%=viskintCHK%> /> <%=erp_txt_240 %>

               <br />
            <input id="Checkbox2" name="FM_viskun_ikkecsv" value="1" type="checkbox" <%=viskicsvCHK%> /> <%=erp_txt_241 %>

             <br /><br />
            <input id="FM_viskun_godkendte" name="FM_viskun_godkendte" value="1" type="checkbox" <%=viskungkCHK%> /> <%=erp_txt_242 %>

            <br />
            <input id="FM_viskun_kladder" name="FM_viskun_kladder" value="1" type="checkbox" <%=viskunklCHK%> /> <%=erp_txt_094 %>

            <br /><br />
            <input id="FM_viskun_fak" name="FM_viskun_fak" value="1" type="checkbox" <%=viskunFakCHK%> /> Fakturaer
            <br />
            <input id="FM_viskun_kre" name="FM_viskun_kre" value="1" type="checkbox" <%=viskunKreCHK%> /> Kreditnotaer




        </td>
	    <%else

      

	    dontshowDD = 1
	    %>
	    <!--#include file="inc/weekselector_s.asp"-->
	    <%end if %>
	    
	 </tr>
	 
	

       
      <%if print <> "j" then  %>
	  <tr>
	  <td valign=top colspan=3>
	    <br /><b><%=erp_txt_227 %></b>&nbsp;<br />
	    <input id="sogFaknrOrJob" name="sogFaknrOrJob" value="0" <%=sogFaknrOrJobCHK0 %> type="radio" /> <%=erp_txt_243 %><br />
        <input id="sogFaknrOrJob" name="sogFaknrOrJob" value="1" <%=sogFaknrOrJobCHK1 %> type="radio" /> <%=erp_txt_244 %><br />
        
        
	    <br /> 
	    <%if print <> "j" then %>
	    <input id="FM_sog" name="FM_sog" value="<%=showSog%>" type="text" style="width:325px;" />&nbsp;<input type="checkbox" value="1" name="FM_ignper" <%=ignperCHK %> /> <%=erp_txt_245 %> <span style="color:#999999;">(søgekri. skal være udfyldt)</span><br />
        <%=erp_txt_246 %><br />
        <%=erp_txt_247 %>
	    <%else %>
	    <%=showSog %>
	    <%end if %>
	    </td>
        </tr>

        <tr>

	    <td align=right valign=bottom colspan=3>
	    <%if print <> "j" then %>
	    <input id="Submit1" type="submit" value="Vis Fakturaer >>" />
	    <%end if %></td>
	 </tr>
	 <%end if %>
	 
	 </form></table>
	 </div>
	 
	 </td></tr>
	 </table>
	 </div>
	
	<%
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	%>
	
	<br />
	<br><b>Periode:</b>&nbsp;
	<%=formatdatetime(strDag&"/"& strMrd &"/"& strAar, 1) & " - " & formatdatetime(strDag_slut &"/"& strMrd_slut &"/"& strAar_slut, 1)%>
	
	<br />
    
    <%
    'Response.Write "jobid: " & jobid
    'Response.Write "<br>aftid: " & aftid 
    
    'Opretter en instans af fil object **'
    Set fs=Server.CreateObject("Scripting.FileSystemObject")



    
    
    tTop = 20
    tLeft = 0
    tWdth = 1204
    
    call tableDiv(tTop,tLeft,tWdth)
    

       
    %>
  <form action="erp_fakhist.asp?func=opdbetalt&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_sog=<%=showSog%>&sort=<%=sort %>" method="post">


    
    <table cellspacing=0 cellpadding=4 border=0 width=100% bgcolor="#ffffff">
  <%if print <> "j" then %>
    <tr><td colspan="20" align="right">  
            <input id="Submit5" type="submit" value="Opdater >>" />
           </td></tr>
         <%end if %>


    <tr bgcolor="#8cAAe6">
        <td class=alt valign=bottom height=25 style="padding:0px 5px 0px 5px;"><b><%=erp_txt_248 %></b><br /><%=erp_txt_206 %></td>
        <!--<td class=alt><b>Job</b></td>
        <td class=alt><b>Aftale</b></td>-->
        <td class=alt valign=bottom colspan=3><b><%=erp_txt_234 %></b></td>
        <td class=alt valign=bottom><b><%=erp_txt_235 %></b><br /><%=erp_txt_249 %><br />
        <%=erp_txt_250 %><br /><%=erp_txt_049 %></td>
        <td class=alt valign=bottom><b><%=erp_txt_007 %></b></td>
        <td class=alt valign=bottom style="padding:0px 0px 0px 2px;"><b><%=erp_txt_093 %></b></td>
        <td class=alt valign=bottom><b><%=erp_txt_251 %></b></td>
         <!--<td class=alt align=right style="padding:0px 5px 0px 0px;"><b>Antal</b></td>-->
        <td class=alt valign=bottom width=100 align=right style="padding:0px 5px 0px 0px;"><b><%=erp_txt_252 %><br /><%=erp_txt_253 %></b></td>
        <td class=alt valign=bottom width=100 align=right style="padding:0px 5px 0px 0px;"><b><%=erp_txt_252 %><br /><%=erp_txt_254 %></b></td>
        
        <td class=alt align=right style="padding:0px 5px 0px 0px;"><b><%=erp_txt_255 %></b><br />
        <%=erp_txt_256 %></td>
         <td class=alt valign=bottom><b><%=erp_txt_257 %></b><br /><b><%=erp_txt_258 %></b></td>
         <td class=alt valign=bottom><b><%=erp_txt_259 %><br /><%=erp_txt_260 %></b></td>
         <td class=alt valign=bottom><b><%=erp_txt_261 %></b></td>
         <td class=alt valign=bottom><b><%=erp_txt_001 %><br /> <%=erp_txt_262 %><br /> <%=erp_txt_263 %></b> <input type="checkbox" id="alleerp" /></td>
         <td class=alt valign=bottom><b><%=erp_txt_264 %></b> <input type="checkbox" id="alleluk" /></td>
         <td>&nbsp;</td>
        
    </tr>
	<%  
	
	
	
	if len(trim(showSog)) <> 0 then
	    
	    if instr(showSog, "-") <> 0 then
	    
	    whereInstr = instr(showSog, "-")
	    showSog1 = mid(showSog, 1, whereInstr-1)
	    showSog1 = trim(showSog1)
	    
	    lenInstr = len(showSog)
	    showSog2 = mid(showSog, whereInstr+1, lenInstr)
	    showSog2 = trim(showSog2)
	    
            if sogFaknrOrJob <> 0 then '** Job: 1 / Faknr: 0
            jobaftSQL = " (j.jobnr BETWEEN "& showSog1 &" AND "& showSog2  &""_
	        &" OR s.aftalenr BETWEEN "& showSog1 &" AND "& showSog2  &") AND shadowcopy <> 1"
            else
            jobaftSQL = " f.faknr BETWEEN "& showSog1 &" AND "& showSog2 &" AND shadowcopy <> 1"
            end if

	   
	    
	    else
	    
            if sogFaknrOrJob <> 0 then '** Job: 1 / Faknr: 0
	        jobaftSQL = " ((j.jobnr LIKE '"& showSog &"' OR j.jobnavn LIKE '"& showSog &"' OR k.kkundenr LIKE '"& showSog &"' OR k.kkundenavn LIKE '"& showSog &"' OR s.aftalenr LIKE '"& showSog &"') AND shadowcopy <> 1) "
	        else
            jobaftSQL = " ((f.faknr LIKE '"& showSog &"' OR f.beloeb = '"& showSog &"') AND shadowcopy <> 1) "
            end if
	    
	    end if
	    
	        %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(showSog) 
			if (isInt > 0 OR instr(showSog, ".") <> 0) AND instr(showSog, "%") = 0 AND instr(showSog, "-") = 0 then
				
		    jobaftSQL = " f.faknr = 9999999 "
			
			isInt = 0
			end if
	
	else
	    
	    if cint(kundeid) <> 0 then
	    kundeIDKri = " AND f.fakadr = " & kundeid 
	    else
	    kundeIDKri = " AND f.fakadr <> 0 "
	    end if
	    
		
		    
		    'if cint(aftid) <> 0 then
		    'jobaftSQL = "(f.aftaleid = "& aftid &")"
		    'else
		    jobaftSQL = "(f.aftaleid <> 0 AND f.jobid = 0 "& kundeIDKri  &") " 
		    'end if
		
		
		
		    'if cint(jobid) = 0 then
		    jobaftSQL = jobaftSQL & " OR (f.jobid <> 0 AND shadowcopy <> 1 "& kundeIDKri &" ) "
		    'else
		    'jobaftSQL = jobaftSQL & " OR ((f.jobid = "& jobid &" AND shadowcopy <> 1) "& kundeIDKri &") "
		    'end if
		    

            if cint(viskint) = 1 then
            intSQLKri = " AND (medregnikkeioms = 1 OR medregnikkeioms = 2) "
            else
            intSQLKri = ""
            end if
	
	
	end if
	    
	    
	      if cint(viskicsv) = 1 then
          overfort_erpSQLKri = " AND overfort_erp = 0"
          else
          overfort_erpSQLKri = ""
          end if

          if cint(viskungk) = 1 then
          kunGodKSQLkri = " AND betalt = 1 "
          else
          kunGodKSQLkri = ""
          end if

          if cint(viskunkl) = 1 then
          kunKLSQLkri = " AND betalt = 0 "
          else
          kunKLSQLkri = ""
          end if


             if cint(viskunFak) = 1 then
             kunFakSQL = " AND faktype = 0"
             else
             kunFakSQL = ""
             end if

             if cint(viskunKre) = 1 then
             kunKreSQL = " AND faktype = 1"
             else
             kunKreSQL = ""
             end if
       
	    
	    
	    kundeidSQL = "k.kid = f.fakadr"
		
		strSQLFak = "SELECT k.kid, f.jobid, f.aftaleid, f.fid, f.fakdato, f.b_dato, f.faknr, f.betalt, "_
		&" f.faktype, f.beloeb, f.timer AS fak, f.shadowcopy, f.erfakbetalt, f.kurs, "_
		&" k.kkundenavn, k.kkundenr, f.brugfakdatolabel, f.istdato2, f.fakbetkom, "_
		&" j.jobnavn, j.jobnr, j.jobslutdato, j.jobstartdato, j.jobstatus, j.fastpris, j.serviceaft, f.istdato, f.istdato2, f.visperiode, j.jobstatus, f.moms, "_
		&" s.navn AS aftnavn, s.advitype , s.aftalenr, s.pris, "_
		&" s.stdato, s.sldato, sj.navn AS sjaftnavn, sj.aftalenr AS sjaftalenr, "
		
	    strSQLFak = strSQLFak &" k1.navn AS konto, k1.kontonr AS kontonr, "_
		&" k2.navn AS modkonto, k2.kontonr AS modkontonr, "_
		&" sum(fms.venter) AS ventetimer, sum(venter_brugt) AS ventetimer_brugt, v.valutakode, v.id AS valid, labeldato, medregnikkeioms, fak_laast, overfort_erp "_
		&" FROM fakturaer f "
		
		strSQLFak = strSQLFak &" LEFT JOIN kunder k on ("&kundeidSQL&")"_
		&" LEFT JOIN job j ON (j.id = f.jobid)"_
		&" LEFT JOIN kontoplan k1 ON (k1.kontonr = f.konto )"_
		&" LEFT JOIN kontoplan k2 ON (k2.kontonr = f.modkonto )"_
		&" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid) "_
	    &" LEFT JOIN serviceaft s ON (s.id = f.aftaleid)"_
	    &" LEFT JOIN serviceaft sj ON (sj.id = j.serviceaft)"_
	    &" LEFT JOIN valutaer v ON (v.id = f.valuta)"
	    
		
	    strSQLFak = strSQLFak &" WHERE ("& jobaftSQL &")"
		
		if len(trim(showSog)) = 0 OR (len(trim(showSog)) <> 0 AND cint(ignper) = 0) then 
		   
		    strSQLFak = strSQLFak &" AND ((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"')"
		    strSQLFak = strSQLFak &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'))"
		    
		    
		    'strSQLFak = strSQLFak &")"
		end if
		
        '** vis kun åbne
		if cint(viskunabne) <> 0 then
		strSQLFak = strSQLFak & " AND erfakbetalt = 0"
		end if

        '** Vis kun dem der ikke er overført til ERP
        strSQLFak = strSQLFak & overfort_erpSQLKri 
	    
        '** vis kun interne
		strSQLFak = strSQLFak & intSQLKri

        '** visk Kun godkendte / Kladder
        strSQLFak = strSQLFak & kunGodKSQLkri & kunKLSQLkri

        '** visk Kun faktuaer / kreditnotaer
        strSQLFak = strSQLFak & kunFakSQL & kunKreSQL

		strSQLFak = strSQLFak &" GROUP BY f.fid ORDER BY "& sortOrderKri &" DESC"
	        
	        'Response.Write strSQLFak &"<br>" 
	        'Response.Flush
	        
		f = 0
		eksportFid = 0
		faksumialt = 0
		fakantalIalt = 0
		belobGrundValTot = 0
		antalAbneposter = 0
		febkbelobAbneposter = 0
        internTotFak = 0
        jobTotFak = 0
        aftTotFak = 0
        kreTotFak = 0
		oRec3.open strSQLFak, oConn, 3
        while not oRec3.EOF
        
        
        
        if oRec3("jobid") <> 0 then 
        bgThis = "#FFFFFF"
        else
        bgThis = "#EFf3ff"
        end if
        
       
            '** Henter Job / aftale oplysninger på faktura ***
            '** A) jobid <> 0 = Faktura oprettet på job.
            
            '** B) aftaleid <> 0 Faktura oprettet på aftale.
            
            '** C) jobid og shadowcopy <> 0 Faktura oprettet på aftale, men
            '** der er et job tilknyttet denne aftale, og derofr bliver der også oprettet en faktura på jobbet.
            
            
            
           
            
            if (oRec3("jobid") <> 0 AND jidaft = 0) OR (oRec3("jobid") <> 0 AND jidaft = 1 AND oRec3("serviceaft") = 0) OR oRec3("aftaleid") <> 0 then 
            
           
            %>
            <tr bgcolor="<%=bgThis %>">
            <td valign=top style="border-bottom:1px #C4C4C4 solid; padding:5px 5px 3px 5px;">
            <span style="color:darkred;"><%=oRec3("kkundenavn") %> (<%=oRec3("kkundenr") %>)</span><br />
            
            
            <%if oRec3("jobid") <> 0 then %>
            
            <u>Job:</u> <b><%=oRec3("jobnavn") %> (<%=oRec3("jobnr") %>)</b>
            <span style="font-size:9px;">
            <%if oRec3("fastpris") = "1" then %>
            <i>- fastpris</i>
            <%else %>
            <i>- lbn. timer</i>
            <%end if %> </span>


                
           
            
               
                  <br /><span style="font-size:9px;">
            <%if isdate(oRec3("jobstartdato")) AND isDate(oRec3("jobslutdato")) then  %>
            <%=replace(formatdatetime(oRec3("jobstartdato"), 2),"-",".") %> til <%=replace(formatdatetime(oRec3("jobslutdato"), 2),"-",".") %>
            <%else %>
            <%=oRec3("jobstartdato") %> til <%=oRec3("jobslutdato") %>
            <%end if %>
            </span>
            
            <%if oRec3("serviceaft") <> 0 then %>
            <br /><font class=medlillesilver><%=oRec3("sjaftnavn") %> (<%=oRec3("sjaftalenr") %>)</font>
            <%end if
           
           
            end if
            
            
            
            '*** Aftaler ***
            if oRec3("aftaleid") <> 0 then
            
            %>
            
            <u>Aftale:</u> <b><%=oRec3("aftnavn") %> (<%=oRec3("aftalenr") %>)</b>
            <font class=megetlillesort>
            <%
             if oRec3("advitype") <> 2 then
	        aftType = "Enh./Klip"
	        else
	        aftType = "Periode"
	        end if
             %>
             
             - <%=aftType %> - 
             
             <%if len(oRec3("stdato")) <> 0 AND len(oRec3("sldato")) <> 0 then %>
             <%=replace(formatdatetime(oRec3("stdato"), 2),"-",".")  %> til <%=replace(formatdatetime(oRec3("sldato"), 2),"-",".")  %>
             <%end if%>
            </font>
            
            <%strSQLaftjob = "SELECT jobnavn, jobnr FROM job WHERE serviceaft = "& oRec3("aftaleid")
            oRec.open strSQLaftjob, oConn, 3
            while not oRec.EOF 
            %>
            <br /><font class=medlillesilver><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)</font>
            <%
            oRec.movenext
            wend
            oRec.close %>
            <%end if %>
            
            
        
        
            
        </td>
        
        
        
        <td valign=top align=right style="border-bottom:1px #C4C4C4 solid;">
        <% if print <> "j" then
            '** hvis godkendt media = 1
            
            if oRec3("jobstatus") = 0 then
            FM_jobonoffval = "j"
            else
            FM_jobonoffval = ""
            end if%>
        <a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec3("fid")%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec3("kid")%>&FM_job=<%=oRec3("jobid")%>&FM_aftale=<%=oRec3("aftaleid")%>&fromfakhist=1" class="vmenu" target="_blank"><%=oRec3("faknr") %></a>
            &nbsp;<a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec3("fid")%>&media=j" target="_blank" class="gray">[Print]</a>&nbsp;
            <!-- erp_fak_godkendt_2007.asp?id=<%=oRec3("fid")%>&aftid=<%=oRec3("aftaleid")%>&jobid=<%=oRec3("jobid")%>" -->
        <%else %>
        <%=oRec3("faknr") %>
        <%end if %>
        
        </td>
         <td valign=top style="border-bottom:1px #C4C4C4 solid; padding-right:5px;">
        <%
        '*** Findes PDF? ***


             if oRec3("faktype") = 1 then
             strFaktypeNavn = "kreditnota"
             else
             strFaktypeNavn = "faktura"
             end if

        strknavn = oRec3("kkundenavn")
        call kundenavnPDF(strknavn)
        strknavn = strKundenavnPDFtxt

        if isNULL(oRec3("labeldato")) <> true then

            if cdate(oRec3("labeldato")) <= cdate("25-08-2011") then
            pdfFakNavn = ""&strFaktypeNavn&"_"&lto&"_"&oRec3("faknr")&".pdf"
            end if

            if cdate(oRec3("labeldato")) > cdate("25-08-2011") AND cdate(oRec3("labeldato")) < cdate("20-09-2011") then
            pdfFakNavn = ""&strFaktypeNavn&"_"&lto&"_"&strknavn&"_"&oRec3("faknr")&".pdf"
            end if

            if cdate(oRec3("labeldato")) >= cdate("20-09-2011") then

            select case lto 
            case "epi_no" 'overflyttet fra epi_sta og epi_osl
                 if cdate(oRec3("labeldato")) >= cdate("25-06-2014") then
                 pdfFakNavn = ""&strFaktypeNavn&"_"&lto&"_"&oRec3("faknr")&"_"&strknavn&".pdf"
                 else
                 pdfFakNavn = ""&strFaktypeNavn&"_epi_osl_"&oRec3("faknr")&"_"&strknavn&".pdf"
                 end if
            case else
            pdfFakNavn = ""&strFaktypeNavn&"_"&lto&"_"&oRec3("faknr")&"_"&strknavn&".pdf"
            end select
            end if

            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\"&toVer&"\timereg\erp_fakhist.asp" then
	            pdfurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&pdfFakNavn
	        else
	            pdfurl = "c:\\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&pdf_FakNavn
	        end if
        
        	    If (fs.FileExists(pdfurl))=true Then
                      %>
                      <a href="https://outzource.dk/timeout_xp/wwwroot/<%=toVer%>/inc/upload/<%=lto%>/<%=pdfFakNavn%>" target="_blank"><img src="../ill/ikon_pdf.gif" border="0"></a>

                        <!--  <a href="http://timeout2.outzource.dk/timeout_xp/wwwroot/<%=toVer%>/inc/upload/<%=lto%>/<%=pdfFakNavn%>" target="_blank"><img src="../ill/ikon_pdf.gif" border="0"></a> -->
                      <%
                Else
                      Response.Write("&nbsp;")
                End If
          

          else

             Response.Write("&nbsp;")

          end if

          %>
          </td>
         <td valign=top style="border-bottom:1px #C4C4C4 solid; padding:4px 5px 0px 0px;">
          <%
        
        '*** Findes XML ***
        if isNULL(oRec3("labeldato")) <> true then
                

                

                 if cdate(oRec3("labeldato")) >= cdate("30-04-2012") then
                 xmlFakNavn = ""&strFaktypeNavn&"_xml_"&lto&"_"&oRec3("faknr")&"_"&strknavn&".xml"
                 else
                 xmlFakNavn = ""&strFaktypeNavn&"_xml_"&lto&"_"&oRec3("faknr")&".xml"
                 end if

                 if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\"&toVer&"\timereg\erp_fakhist.asp" then
	                xmlurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&xmlFakNavn
	            else
	                xmlurl = "c:\\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&xmlFakNavn
	            end if
        
        	        If (fs.FileExists(xmlurl))=true Then

                           if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                            Response.Write("<a href=""https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/inc/upload/"&lto&"/"&xmlFakNavn&""" target=""blank""><img src=""../ill/ikon_xml.gif"" border=""0""></a>")
                          else
                            Response.Write("<a href=""http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/inc/upload/"&lto&"/"&xmlFakNavn&""" target=""blank""><img src=""../ill/ikon_xml.gif"" border=""0""></a>")
                           end if
                    Else
                          Response.Write("&nbsp;")
                    End If
        

         else

             Response.Write("&nbsp;")

          end if
       
        
        %>
        
        
        <%if oRec3("jobid") <> 0 then
        thisJobId = oRec3("jobid")
        else
        thisJobId = 0
        end if
        
        if oRec3("aftaleid") <> 0 then
        thisAftaleId = oRec3("aftaleid")
        else
        thisAftaleId = 0
        end if
         %>
         
           
           
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid; white-space:nowrap;">
        
        <%if cint(oRec3("brugfakdatolabel")) = 1 then %>
        L: <b><%=replace(formatdatetime(oRec3("labeldato"),2),"-",".")  %></b><br />
        <span style="font-size:9px; color:#999999;">(<%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %>)</span>
        <%else %>
        F: <b><%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %></b>
        <%end if %>

       
        <%if oRec3("visperiode") = 1 then %>
         <br />
        <span style="font-size:8px; color:#5582d2;"><%=replace(formatdatetime(oRec3("istdato"),2),"-",".") %> - <%=replace(formatdatetime(oRec3("istdato2"),2),"-",".") %></span>
        <%else %>
           <br />
        <span style="font-size:8px; color:#999999;">(<%=replace(formatdatetime(oRec3("istdato"),2),"-",".") %> - <%=replace(formatdatetime(oRec3("istdato2"),2),"-",".") %>)</span>
        <%end if %>
        

        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid;">
        <%if len(oRec3("b_dato")) <> 0 then %>
        <%=replace(formatdatetime(oRec3("b_dato"),2),"-",".") %>
        <%else %>
        -
        <%end if %>
        
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid; padding:2px 5px 0px 5px; white-space:nowrap;">
        <%if oRec3("betalt") = 1 then %>
       <a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec3("fid")%>&FM_kunde=<%=oRec3("kid")%>&FM_job=<%=thisJobId%>&FM_aftale=<%=thisAftaleId%>" class="vmenu" target="_blank"><img src="../ill/godkend_icon_16.gif" border="0" /> <%=erp_txt_095 %></a>
        <%else %>
        
         <%if cint(oRec3("betalt")) = 0 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) AND print <> "j" then %>
           <a href="../timereg/erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&visfaktura=1&func=red&id=<%=oRec3("fid")%>&FM_kunde=<%=oRec3("kid")%>&FM_job=<%=thisJobId%>&FM_aftale=<%=thisAftaleId%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&reset=1" class="gray" target="_blank">
               <!-- &FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>-->
           <img src="../ill/redigerfak.gif" border="0" alt="Rediger faktura" /> <%=erp_txt_094 %></a>

        <%else %>
            <font color=silver><%=erp_txt_094 %></font>
           <%end if %>
        
        
        <%end if %>
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid;">
        <%select case oRec3("faktype")
        case 1
        strType = "Kreditnota"
        minus = "-"
        case else
        strType = "Faktura"
        minus = ""
        end select
        
        Response.Write strType
        %>

          <%if cint(oRec3("medregnikkeioms")) = 1 then %>
         <br /><%=erp_txt_265 %>
        <%end if %>

         <%if cint(oRec3("medregnikkeioms")) = 2 then %>
         <br /><%=erp_txt_266 %>
        <%end if %>

        </td>
        <!--
        <td valign=top align=right style="border-bottom:1px #C4C4C4 solid; padding:1px 5px 0px 0px;">
        <formatnumber(minus&oRec3("fak"), 2)%>
        </td>
        -->
       
        <td valign=top align=right style="border-bottom:1px #C4C4C4 solid; padding:1px 5px 0px 0px;"> 

        <%fakbelob = replace(oRec3("beloeb"), "-", "") %>

        <b><%=formatnumber(minus&fakbelob, 2)&" "&oRec3("valutakode") %></b>
      
        <br />
              
        <%
        '*** Beløb i grundvaluta == DKK ***'
        if oRec3("valid") <> 1 then 
            
            call beregnValuta(minus&(fakbelob),oRec3("kurs"),100)
            belobGrundVal = valBelobBeregnet %>
            
            ~ <%=formatnumber(belobGrundVal, 2) &" "& basisValISO %>
        
            <% 
        else
            
            belobGrundVal = minus&fakbelob/1
        end if 

        '** job ell. aftale **'
        
        



        if oRec3("faktype") = 1 then '** kreditnota
        kreTotFak = kreTotFak/1 + belobGrundVal
        else

            '** kreditnotaer skal ikke medregnes ***'
            if oRec3("jobid") <> 0 then '** jobfak
            jobTotFak = jobTotFak/1 + belobGrundVal 
            else
            aftTotFak = aftTotFak/1 + belobGrundVal
            end if

            if oRec3("medregnikkeioms") <> 0 then
            internTotFak = internTotFak/1 + belobGrundVal/1
            else
            belobGrundValTot = belobGrundValTot/1 + belobGrundVal/1 
            end if

            

        end if
        
        
        '*** Åbne poster ****'
        if cint(oRec3("erfakbetalt")) = 1 then
        febkbelobAbneposter = febkbelobAbneposter 
        else
            if oRec3("medregnikkeioms") = 0 then
            febkbelobAbneposter = febkbelobAbneposter + (belobGrundVal/1)  
            antalAbneposter = antalAbneposter + 1
            else
            febkbelobAbneposter = febkbelobAbneposter 
            end if      
        end if%>
        
       
        </td>

        <td valign=top align=right style="border-bottom:1px #C4C4C4 solid; padding:1px 5px 0px 0px; color:#999999;"> 

        <%fakbelobInklMoms = fakbelob/1 + replace(oRec3("moms")/1, "-", "") %>

        <b><%=formatnumber(minus&fakbelobInklMoms, 2)&" "&oRec3("valutakode") %></b>
        </td>
         
        <td valign=top align=right style="border-bottom:1px #C4C4C4 solid; padding:1px 5px 0px 0px;"><%=oRec3("ventetimer_brugt")%> / <%=oRec3("ventetimer") %></td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid; padding:2px 5px 0px 0px;" class=lillegray><%=oRec3("konto") %>&nbsp;<%=oRec3("kontonr") %><br /><%=oRec3("modkonto") %>&nbsp;<%=oRec3("modkontonr") %></td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid; padding:3px 5px 0px 0px;">
            <%if oRec3("erfakbetalt") = 1 then
            efbCHK = "SELECTED"
            ikbCHK = ""
            else
            ikbCHK = "SELECTED"
            efbCHK = ""
            end if 
            
            if print <> "j" then%>
            <!--
            <input id="fakid_<%=oRec3("fid")%>" name="fakid_<%=oRec3("fid")%>" value="1" type="radio" <%=efbCHK %> /> <b>Ja</b><br />
            <input id="fakid_<%=oRec3("fid")%>" name="fakid_<%=oRec3("fid")%>" value="0" type="radio" <%=ikbCHK %> /> Nej
            -->
            <select id="fakid_<%=oRec3("fid")%>" name="fakid_<%=oRec3("fid")%>" style="width:50px; font-size:9px;">
                <option value="0" <%=ikbCHK %>><%=erp_txt_267 %></option>
                <option value="1" <%=efbCHK %>><%=erp_txt_268 %></option>
            </select>
            <input id="erfakbetalt" name="erfakbetalt" type="hidden" value="<%=oRec3("fid")%>" />
           
            <%else 
                if oRec3("erfakbetalt") = 1 then
                Response.Write "<i>V</i>"
                else
                Response.Write "-"
                end if
            end if
            %>
        </td>
        
        <td valign=top class=lille style="border-bottom:1px #C4C4C4 solid;">
		<%if print <> "j" then %>
		<textarea id="FM_fakbetkom" name="FM_fakbetkom" cols="20" rows="3" style="width:133px; font-size:9px; font-family:Arial;" ><%=oRec3("fakbetkom") %></textarea>	
            <input id="FM_fakbetkom" name="FM_fakbetkom" value="#" type="hidden" />	
            <%else %>
            <%=oRec3("fakbetkom") %>&nbsp;
            <%end if %>		
		</td>

        <%if oRec3("overfort_erp") = 1 then
        overfort_erpCHK = "CHECKED"
        else
        overfort_erpCHK = ""
        end if %>
        
        <td valign=top style="border-bottom:1px #C4C4C4 solid;"><input type="checkbox" name="overfortERP" value="1" class="overfortERP" <%=overfort_erpCHK %> /><input type="hidden" value=";" name="overfortERP" /></td>
        <td valign=top style="border-bottom:1px #C4C4C4 solid;">


        <%if thisJobId <> 0 AND oRec3("jobstatus") <> 0 then %>
        <input type="checkbox" name="lukjob" value="<%=thisJobId %>" class="lukjob" /><input type="hidden" value=";" name="lukjob" />
        <%else 
                 %>

                
                <span style="font-size:9px; color:#999999; padding:1px; border:1px silver solid;"><%=erp_txt_269 %></span>
               

            <input type="hidden" value="0" name="lukjob" /><input type="hidden" value=";" name="lukjob" />
        &nbsp;
        <%end if %>
        </td>

        <td valign=top style="border-bottom:1px #C4C4C4 solid;">
            <%if cint(oRec3("betalt")) = 0 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) AND print <> "j" AND oRec3("fak_laast") = 0 then %>
            <a href="erp_opr_faktura_fs.asp?func=slet&rdir=hist&id=<%=oRec3("fid")%>&visminihistorik=0&visfaktura=0&visjobogaftaler=0&nomenu=1" target="_blank"><img src="../ill/slet_16.gif" border=0 /></a>
            <!-- javascript:popUp(' erp_fak.asp?func=slet&rdir=hist&id=<%=oRec3("fid")%> -->
            <%else %>
            &nbsp;
            <%end if %>
            </td>
        </tr>
      
        <% 
        
        'if oRec3("jobid") <> 0 AND oRec3("serviceaft") <> 0 then
        'fakantalIalt = fakantalIalt 
        'else 
        'fakantalIalt = fakantalIalt + (minus&oRec3("fak"))
        'end if
        
        Response.flush
        f = f + 1
        eksportFid = eksportFid &","& oRec3("fid")
        
        end if 
        
        oRec3.movenext
        wend
        oRec3.close
	
	
	
	
	if f = 0 then%>
	<tr><td colspan=14><%=erp_txt_270 %></td></tr>
	</table>
	<%else %>
	<tr><td>
        &nbsp;</td>
        <td colspan=3 style="white-space:nowrap;">&nbsp;</td>
     
        <!--<td align=right style="padding:0px 5px 0px 0px;">&nbsp;</td>-->
        <td colspan=5 align=right style="padding:10px 5px 0px 0px;">
        <%=erp_txt_271 %> <b><%=f %></b> <%=erp_txt_036 %><br /><br />
        <%=erp_txt_272 %> <B><%=formatnumber((belobGrundValTot+internTotFak+(kreTotFak)), 2) &" "&basisValISO  %></B> <br />
        <%=erp_txt_273 & " " & formatnumber(kreTotFak, 2) & " "& basisValISO %>)<br /><br />
        <%=erp_txt_274 & " " & formatnumber(internTotFak, 2) & " "& basisValISO %><br />
        <%=erp_txt_280 %> <B><%=formatnumber(belobGrundValTot+(kreTotFak), 2) &" "&basisValISO  %></B> <br /><br />
        
        <%=erp_txt_275 &" "& formatnumber(jobTotFak, 2) & " "& basisValISO %><br />
        <%=erp_txt_276 &" "& formatnumber(aftTotFak, 2) & " "& basisValISO %><br /><br />
        
        <%=erp_txt_277 %> <b><%=antalAbneposter %></b> <%=erp_txt_278 &" "& formatnumber(febkbelobAbneposter, 2) &" "&basisValISO  %>
        <br /><br />
        <%=erp_txt_279 %>
        </td>
        
        
        <td colspan=10 align=right valign=top><br />
        <%if print <> "j" then %>
            <input id="Submit1" type="submit" value="Opdater >>" />
            <%end if %>
        &nbsp;</td>
    </tr>
     <input id="FM_fakbetkom" name="FM_fakbetkom" value="xc" type="hidden" />
     <input type="hidden" value="0" name="overfortERP" />
     <input type="hidden" value="0" name="lukjob" />	
   
    </table>
       </form>
    <br /><br /><br />&nbsp;

    </div>
    
    
    
    
    <%
    
    if print <> "j" then
    
    itop = 80
    ileft = 0
    iwdt = 400
    
    call sideinfo(itop,ileft,iwdt)
    %>
    <%=erp_txt_281 %>
 
   </td>
    </tr>
    </table>
    </div>
    <br /><br /> <br /><br /><br /> <br /><br /><br />
    <br /><br /><br />
    &nbsp;
   <%
        
        
       ptop = 0
       pleft = 875 
       pwdt = 250 
        
       call eksportogprint(ptop,pleft,pwdt)
        
        %>
        
      
       <form action="erp_fakturaer_eksport_2007.asp?visning=0" method="post" target="_blank">
     <input id="Hidden2" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit4" type="submit" value="A) Detail .csv fil eksport" style="font-size:9px; width:160px;" />
    <!--<br /><font class=megetlillesort><%=erp_txt_282 %>-->
    </td>
    </tr>
    </form>

    
      <form action="erp_fakturaer_eksport_2007.asp?visning=1" method="post" target="_blank">
     <input id="fakids" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit2" type="submit" value="B) Standard .csv fil eksport." style="font-size:9px; width:160px;" />
        <!--<br /><font class=megetlillesort><%=erp_txt_283 %>-->
    </td>
    </tr>
    </form>

        <form action="erp_fakturaer_eksport_2007.asp?visning=2" method="post" target="_blank">
     <input id="Hidden1" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit3" type="submit" value="C) Minimun .txt fil eksport." style="font-size:9px; width:160px;" />
        <!--<br /><font class=megetlillesort><%=erp_txt_284 %>-->
    </td>
    </tr>
    </form>

   
     <tr>
    <td align=center><a href="erp_fakhist.asp?print=j&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_sog=<%=showSog%>&FM_viskunabne=<%=viskunabne%>&sort=<%=sort %>&FM_viskun_interne=<%=viskint%>" class=rmenu target=blank><img src="../ill/printer3.png" border=0 alt="" /></a>
    </td><td><a href="erp_fakhist.asp?print=j&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_sog=<%=showSog%>&FM_viskunabne=<%=viskunabne%>&sort=<%=sort %>&FM_viskun_interne=<%=viskint%>" class=rmenu target=blank><%=erp_txt_285 %></a></td>
   </tr>
   
   <tr>
    <td align=center valign=top><a href="#" onclick="Javascript:window.open('erp_make_pdf_multi.asp?fakids=<%=eksportFid%>&lto=<%=lto%>&doctype=pdfxml', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/ikon_sendemail_24.png" border=0 alt="" /></a>
    </td><td><a href="#" onclick="Javascript:window.open('erp_make_pdf_multi.asp?fakids=<%=eksportFid%>&lto=<%=lto%>&doctype=pdfxml', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><%=erp_txt_286 %></a>
    <font class=megetlillesort><%=erp_txt_287 %>
    <br />&nbsp;</td>
   </tr>
   
 
   
     </table>
        </div>
    <%else %>   
     <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %> 
        
    <%end if%>
	
	
	
	
	<%end if 'f = 0%>
	
	<%


    set fs=nothing
	
	end select 'func %>
                
         
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->