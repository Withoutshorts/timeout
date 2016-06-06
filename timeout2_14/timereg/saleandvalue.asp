<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
    <!--#include file="../inc/regular/topmenu_inc.asp"-->






<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else

    '**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
	
    menu = request("menu")
    thisfile = "saleandvalue"

    media = request("media")
    if media = "print" OR media = "export" then
        print = "j"
    end if

	func = request("func")
	id = 0
	
    if len(trim(request("FM_kunde"))) <> 0 AND request("FM_kunde") <> 0 then
    kundeid = request("FM_kunde")
    else
    kundeid = 0
    end if


    '**** Kolonner ****'
    if len(trim(request("FM_kunde"))) <> 0 then
        if len(trim(Request("FM_kol_1"))) <> 0 AND Request("FM_kol_1") <> 0 then
        viskol1 = 1
        viskol1CHK = "CHECKED"
        else
        viskol1 = 0
        viskol1CHK = ""
        end if
    else
        if request.cookies("stat")("viskol1") <> "" then
        viskol1 = 1
        viskol1CHK = "CHECKED"
        else
        viskol1 = 0
        viskol1CHK = ""
        end if

    end if


    response.cookies("stat")("viskol1") = viskol1


     if len(trim(request("FM_kunde"))) <> 0 then
        if len(trim(Request("FM_kol_2"))) <> 0 AND Request("FM_kol_2") <> 0 then
        viskol2 = 1
        viskol2CHK = "CHECKED"
        else
        viskol2 = 0
        viskol2CHK = ""
        end if
    else
        if request.cookies("stat")("viskol2") <> "" then
        viskol2 = 1
        viskol2CHK = "CHECKED"
        else
        viskol2 = 0
        viskol2CHK = ""
        end if

    end if


    response.cookies("stat")("viskol2") = viskol2


     if len(trim(request("FM_kunde"))) <> 0 then
        if len(trim(Request("FM_kol_3"))) <> 0 AND Request("FM_kol_3") <> 0 then
        viskol3 = 1
        viskol3CHK = "CHECKED"
        else
        viskol3 = 0
        viskol3CHK = ""
        end if
    else
        if request.cookies("stat")("viskol3") <> "" then
        viskol3 = 1
        viskol3CHK = "CHECKED"
        else
        viskol3 = 0
        viskol3CHK = ""
        end if

    end if


    response.cookies("stat")("viskol3") = viskol3


     if len(trim(request("FM_kunde"))) <> 0 then
        if len(trim(Request("FM_kol_4"))) <> 0 AND Request("FM_kol_4") <> 0 then
        viskol4 = 1
        viskol4CHK = "CHECKED"
        else
        viskol4 = 0
        viskol4CHK = ""
        end if
    else
        if request.cookies("stat")("viskol4") <> "" then
        viskol4 = 1
        viskol4CHK = "CHECKED"
        else
        viskol4 = 0
        viskol4CHK = ""
        end if

    end if


    response.cookies("stat")("viskol4") = viskol4


     if len(trim(request("FM_kunde"))) <> 0 then
        if len(trim(Request("FM_kol_5"))) <> 0 AND Request("FM_kol_5") <> 0 then
        viskol5 = 1
        viskol5CHK = "CHECKED"
        else
        viskol5 = 0
        viskol5CHK = ""
        end if
    else
        if request.cookies("stat")("viskol5") <> "" then
        viskol5 = 1
        viskol5CHK = "CHECKED"
        else
        viskol5 = 0
        viskol5CHK = ""
        end if

    end if


    response.cookies("stat")("viskol5") = viskol5
	

   


     if len(trim(request("FM_kunde"))) <> 0 then
        if len(trim(Request("FM_kol_6"))) <> 0 AND Request("FM_kol_6") <> 0 then
        viskol6 = 1
        viskol6CHK = "CHECKED"
        else
        viskol6 = 0
        viskol6CHK = ""
        end if
    else
        if request.cookies("stat")("viskol6") <> "" then
        viskol6 = 1
        viskol6CHK = "CHECKED"
        else
        viskol6 = 0
        viskol6CHK = ""
        end if

    end if

     response.cookies("stat")("viskol6") = viskol6


	'*** År ***'
		if len(Request("seomsfor")) <> 0 then
		seomsfor = Request("seomsfor")
		strYear = seomsfor
		strReqAar = strYear 
		else
			
			seomsfor = year(now)
			strYear = seomsfor
			strReqAar = strYear
			
		end if
	
    if len(trim(request("FM_start_mrd"))) <> 0 then
	strMrd = request("FM_start_mrd")
	else
	strMrd = month(now)
	end if
	

    stDato = "1/"& strMrd &"/"& strYear 

                       
                        
                      
	if len(trim(request("FM_virksomhedsandel"))) <> 0 AND request("FM_virksomhedsandel") <> 0 then
    virkAndelChk = "CHECKED"
    virkAndel = 1
    else
    virkAndelChk = ""
    virkAndel = 0
    end if
	
		
	
		
	
	'*** Medarbejdere / projektgrupper
	selmedarb = session("mid")
	'call medarbogprogrp("oms")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAnsSQLkri = ""
	jobAns2SQLkri = ""
	fakmedspecSQLkri = ""
	
	    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
        if request.cookies("tsa")("pgrp") <> "" then
	    progrp = request.cookies("tsa")("pgrp")
        else
        progrp = 0
	    end if
    end if

    response.cookies("tsa")("pgrp") = progrp
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then
	        if media <> "print" then
	        thisMiduse = request("FM_medarb_hidden")
    	    else
    	    thisMiduse = request("FM_medarb")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
	

	
	
	'*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			     'kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     'jobAnsSQLkri = "jobans1 = "& intMids(m)  
			     'jobAns2SQLkri = "jobans2 = "& intMids(m)
                 'jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     'fakmedspecSQLkri = "fms.mid = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     'kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     'jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     'jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
                 'jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     'fakmedspecSQLkri = fakmedspecSQLkri & " OR fms.mid = "& intMids(m)
			     end if
			     
			    next
			    
			    'medarbSQlKri = medarbSQlKri & ")"
			    
			jobAnsSQLkri =  " ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "xx (" & jobAns2SQLkri &")"
			jobAns3SQLkri =  "xx (" & jobAns3SQLkri &")"
            
            'fakmedspecSQLkri = " AND ("& fakmedspecSQLkri &")"
			
	
	
	

     if len(trim(request("interval"))) <> 0 then
	 interval = request("interval")
     end if
     
     'Response.write "interval " & interval

     
             
     select case interval 
	 case "1"
     slDato = dateAdd("m", 1, stDato)
     slDato = dateAdd("d", -1, slDato)
     intv_1CHK = "CHECKED"
     stDato = dateAdd("m", 0, stDato)
     case "3"
     slDato = dateAdd("m", 1, stDato)
     slDato = dateAdd("d", -1, slDato)
     intv_3CHK = "CHECKED"
     stDato = dateAdd("m", -2, stDato)
     case "6"
     slDato = dateAdd("m", 1, stDato)
     slDato = dateAdd("d", -1, slDato)
     intv_6CHK = "CHECKED"
     stDato = dateAdd("m", -5, stDato)
     case "12"
     slDato = dateAdd("m", 1, stDato)
     slDato = dateAdd("d", -1, slDato)
     intv_12CHK = "CHECKED"
     stDato = dateAdd("m", -11, stDato)
     case else
     slDato = dateAdd("m", 1, stDato)
     slDato = dateAdd("d", -1, slDato)
     interval = 3
     intv_3CHK = "CHECKED"
     end select
	
     stDatoShow = stDato
     slDatoShow = slDato
     stDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato) 
	 slDato = year(slDato) &"/"& month(slDato) &"/"& day(slDato)
	 
     if len(trim(request("FM_toplinie_db"))) <> 0 then 
	 toplinie_db = request("FM_toplinie_db")
     else
     toplinie_db = 0 '= DB
     end if
	
    

    if cint(toplinie_db) = 0 then 'DB
    toplinie_dbCHK0 = "CHECKED"
    toplinie_dbCHK1 = ""
    toplinie_txt = "DB."
    else
    toplinie_dbCHK0 = ""
    toplinie_dbCHK1 = "CHECKED"
    toplinie_txt = "Toplinie"
    end if

	'************ slut faste filter var **************		
	

	if media = "print" OR media = "export" OR media = "chart" then
	'leftPos = 20
	'topPos = 62
    twth = 95
    %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	else
	'leftPos = 90
	'topPos = 102
    twth = 100
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

	
<!--
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%
        'call tsamainmenu(7)
    %>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
        'call stattopmenu()
    %>
	</div>-->

	<%

        call menu_2014()
	end if
	%>


     <script src="inc/sale_jav_load.js"></script>

    <!--#include file="inc/dato2_b.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	        

            <%if media <> "print" AND media <> "export" AND media <> "chart" then
            lbTop = 360
            lbLeft = 300
            else
            lbTop = 160
            lbLeft = 200
            end if%>
	        
            <div id="loadbar" style="position:absolute; display:; visibility:visible; top:<%=lbTop%>px; left:<%=lbleft%>px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	        <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	        <img src="../ill/outzource_logo_200.gif" />
	        <br />
	        Forventet loadtid:
	        <%
	
	        exp_loadtid = 0
	        'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	        ca.: <b>3-35 sek. alt efter antal medarb. og kolonne valg.</b>
	        </td><td align=right style="padding-right:40px;">
	        <img src="../inc/jquery/images/ajax-loader.gif" />
	
	        </td></tr></table>

	        </div>
	
            

	
	<%
	if media <> "print" AND media <> "export" AND media <> "chart" then
	pleft = 90
	ptop = 102
	'ptopgrafik = 348
	else
	pleft = 10
	ptop = 20
	'ptopgrafik = 90
	end if

    if media <> "export" then
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
	
	
	<%
	
	

	
	if media <> "print" AND media <> "export" AND media <> "chart" then 
	
	call filterheader_2013(0,0,800,pTxt)
        
        

        'oimg = "ikon_sale_48.png"
	oleft = 0
	otop = 0
	owdt = 650
	
    if cint(virkAndel) = 1 then
    virkAndelTxt = "<b>Virksomhedsandel medregnet</b>,"
    else
    virkAndelTxt = ""
    end if

    oskrift = "Salg og Værdi pr. medarbejder<br><span style=""font-size:11px; line-height:12px; font-style:normal; font-weight:lighter;""><br>Viser alle job hvor valgte medarbejdere er med som jobansvarlig, jobejer, eller job-medansvarlig 1-3. og hvor deres salgsandel er <> 0 % (maks. 100% andel). "&virkAndelTxt&" DB ell. toplinie: <b>"& toplinie_txt &"</b> </span>"
	


	call sideoverskrift_2014(oleft, otop, owdt, oskrift)

    end if

        
        
        
        
        
        %>
    	    <br />
	  
	    <table cellspacing=0 cellpadding=10 border=0 width=100% bgcolor="#FFFFFF">
        
	    <form action="saleandvalue.asp?rdir=<%=rdir %>" method="post">
	    
	    
        
        <tr>
   
	    
	    <%call progrpmedarb %>
	    
	
	</tr> 
    <tr>
    <%
    kundeAnsSQLKri = " (Kundeans1 <> -1 AND kundeans2 <> -1) "
    call segment_kunder  %>
    </td>
    <td style="border-bottom:1px solid #cccccc;">&nbsp;</td>
    
    </tr>
    
	<tr>
		<td valign=top>
        <br />
        <b>Vis job / fakturaer med job-startdato / fakturadato
        fra:</b><br /><br />
	<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
	&nbsp;
        
     <b>År:</b>&nbsp;
	<select name="seomsfor" id="seomsfor" style="width:85px;">
			
			<%
			useaar = 2001
			for x = 0 to 20
			useaar = useaar + 1
			
			if cint(strYear) = cint(useaar) then
			aChk = "SELECTED"
			else
			aChk = ""
			end if
			%>
			<option value="<%=useaar%>" <%=aChk%>><%=useaar%></option>
			<%
			'x = x + 1
			next%>
		</select>
        <br /><br />

       Vælg periode, fra viste måned og forrige<br /> <input type="radio" name="interval" value="12" <%=intv_12CHK %> />12 md. 
       <input type="radio" name="interval" value="6" <%=intv_6CHK %> />6 md.
       <input type="radio" name="interval" value="3" <%=intv_3CHK %> />kvartal
       <input type="radio" name="interval" value="1" <%=intv_1CHK %> />kun valgte måned.
       
     
       </td>
       <td valign="top"><br />
       <b>DB ell. toplinie:</b><br />
       Vis som <input id="Radio2" name="FM_toplinie_db" value="0" type="radio" <%=toplinie_dbCHK0 %> />  <b>DB</b>  eller <input id="Radio3" name="FM_toplinie_db" value="1" <%=toplinie_dbCHK1 %> type="radio" /> <b>Toplinie</b> (før salgsomkostninger, Tilbud & Job (WIP) budget)
       
       <%if level = 1 then %>
       <br /><br />
       <b>Virksomhedsandel:</b><br />
       <input type=checkbox value=1 name="FM_virksomhedsandel" <%=virkAndelCHK %> /> Medregn virksomhedsandel angivet på hvert job.
       <%end if %>

       <br /><br />

        <b>Vis kolonner:</b><br />
       <input type=checkbox value=1 name="FM_kol_1" <%=viskol1CHK %> />Tilbud & Job (WIP) budget<br />
       <input type=checkbox value=1 name="FM_kol_2" <%=viskol2CHK %> />Realiseret og faktureret, lukkede job<br />
       <input type=checkbox value=1 name="FM_kol_3" <%=viskol3CHK %> />Realiseret og faktureret, åbne job<br />
       <!--<input type=checkbox value=1 name="FM_kol_5" <%=viskol5CHK %> />Salgsomkostninger<br />-->
       <input type=checkbox value=1 name="FM_kol_6" <%=viskol6CHK %> />Effektiv timepris (ALLE lukkede job i periode uanset andel)<br />
       <input type=checkbox value=1 name="FM_kol_4" <%=viskol4CHK %> />Grundlag<br />


	</td>
    </tr>
    <tr>
	<td colspan="2" align=right valign=bottom>
	&nbsp;<input type="submit" value=" Kør >> ">
	</td></tr>
    </form>
	</table>

   
	
	<!--filter div-->
	</td></tr>
	</table>
	</div>
    <%else 
    
     kundeAnsSQLKri = " (Kundeans1 <> -1 AND kundeans2 <> -1) "
    call segment_kunder 

    end if %>

	
	
	
	
	
        <%
	   


	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	

    if media <> "export" AND media <> "chart" then

    %>
    <br /><br />
    <b>&nbsp;Periode: <%=formatdatetime(stDatoShow, 1) %></b> til <b><%=formatdatetime(slDatoShow, 1) %></b>

    <%
	
    tTop = 10
	tLeft = 0
	tWdth = 1300
	
	
	call tableDiv(tTop,tLeft,tWdth)

    end if

    
    public thisProc
    function jobansproc(jobans_flt)

        if jobans_flt <> 0 then
        thisProc = jobans_flt
        else
        thisProc = 0
        end if

    end function
    

    strExport = ""

    if media <> "export" AND media <> "chart" then %>
	
    <table cellpadding=2 cellspacing=1 border=0 width="<%=twth%>%" bgcolor="#CCCCCC">
    
     <tr bgcolor="#5582d2">
    <td bgcolor="#FFFFFF" valign=bottom><b>Medarbejder</b></td>
    <%if cint(viskol1) = 1 then %>
    <td class=alt colspan=2 bgcolor="#8cAAe6" valign=bottom><b>Tilbud : Andel % (jobansvarlig)</b></td>
    <td class=alt bgcolor="#5C75AA" valign=bottom><b>Aktive job : WIP* (restestimat) : Andel % (jobansvarlig)</b></td>
    <%end if %>
    
  

       <%if cint(viskol2) = 1 then %>
         <td class=alt bgcolor="#999999" colspan=7 valign=bottom><b>Lukkede job i perioden : Andel % (jobansvarlig)</td>
       <%end if %>

          <%if cint(viskol3) = 1 then %>
      <td class=alt bgcolor="#5C75AA" colspan=6 valign=bottom><b>Alle åbne job, med slutdato senere end periode startdato (<%=formatdatetime(stDatoShow, 1) %>) : Andel % (jobansvarlig)</td>
     <%end if %>

     

      <%if cint(viskol6) = 1 then %>
   
      <td bgcolor="#FFFFFF" valign=bottom><b>Effektiv timepris</b><br />Alle lukkede job i periode, uanset andel</td>
     <%end if %>
 
    <%if cint(viskol4) = 1 then %>
    <td class=alt valign=bottom><b>Grundlag</b></td>
    <%end if %>
    </tr>

    <tr bgcolor="#5582d2">
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <%if cint(viskol1) = 1 then %>
    <td class=alt valign=bottom bgcolor="#8cAAe6"><b>Tilbud : Andel % : <%=toplinie_txt &" "& basisValISO %></b></td>
    <td class=alt valign=bottom bgcolor="#8cAAe6"><b>Tilbud * Sandsyn. % : Andel :  <%=toplinie_txt &" "& basisValISO %></b></td>
    <td class=alt valign=bottom bgcolor="#5C75AA"><b>Aktive job : WIP* (restestimat) : Andel : <%=toplinie_txt &" "& basisValISO %></b></td>
    <%end if %>
    
    <%if cint(viskol2) = 1 then %>
     <td class=alt valign=bottom bgcolor="#999999"><b>Real. timer</b></td>
    <td class=alt valign=bottom bgcolor="#999999"><b>Real. omsætning. <%=basisValISO %></b><br /><span style="font-size:9px;">(timepris * timer)</span></td>
    <td class=alt valign=bottom bgcolor="#999999"><b>Real. omkost. <%=basisValISO %></b><br /><span style="font-size:9px;">(intern kost.)</span></td>
    <td class=alt valign=bottom bgcolor="Yellowgreen"><b>Faktureret oms. <%=basisValISO %></b></td>
    <td class=alt valign=bottom bgcolor="#999999"><b>Sum <%=basisValISO %></b><br />
    <span style="font-size:9px;">(faktureret - intern kost.)</span></td>
    <td class=alt valign=bottom bgcolor="#FFCC66"><b>Salgsomk. lukkede job <%=basisValISO %></b></td>
    <td class=alt bgcolor="#999999" valign=bottom><b>Faktisk timepris <%=basisValISO %></b><br />
    <span style="font-size:9px;">
    (fakturet beløb - salgsomk. / realiseret timer) 
    </span></td>
     
      <%end if %>

      <%if cint(viskol3) = 1 then %>
<td class=alt bgcolor="#5C75AA" valign=bottom><b>Real. timer</b></td>
    <td class=alt bgcolor="#5C75AA" valign=bottom><b>Real. omkost. <%=basisValISO %></b><br /><span style="font-size:9px;">(intern kost.)</span></td>
    <td class=alt bgcolor="Yellowgreen" valign=bottom><b>Faktureret oms. <%=basisValISO %></b></td>
    <td class=alt bgcolor="#5C75AA" valign=bottom><b>Sum <%=basisValISO %></b><br /><span style="font-size:9px;">(faktureret - intern kost.)</span></td>
    <td class=alt valign=bottom bgcolor="#FFCC66"><b>Salgsomk. åbne job <%=basisValISO %></b></td>
     <td class=alt bgcolor="#5C75AA" valign=bottom><b>Faktisk timepris <%=basisValISO %></b><br />
    <span style="font-size:9px;">
    (fakturet beløb - salgsomk. / realiseret timer) 
    </span></td>
    <%end if %>
    
       

    <%if cint(viskol6) = 1 then %>
     <td bgcolor="#FFFFFF" valign=bottom><b>Faktor</b><br />
      <span style="font-size:9px;">(kvotient, egne fakturerbare timer vægtet ifht. <br />
      faktureret - salgsomk. / kostpr. pr .job)</span></td>
      <%end if %>
    
    <%if cint(viskol4) = 1 then %>
    <td class=alt>&nbsp;</td>
    <%end if %>
    </tr>
    
    <%
    end if
    
    dim strExportA, medarbBelob, medarbGrlag, jobidsEgneTimer, jobidsSalgsomkost, jobidsFaktureret
    redim strExportA(1950,15), medarbBelob(1950,15), medarbGrlag(1950,15), jobidsEgneTimer(20000), jobidsSalgsomkost(20000), jobidsFaktureret(20000)


    '** medarbBelob 1 = Tilbud
    '** medarbBelob 2 = Tilbud
    '** medarbBelob 3 = Job / WIP

    '** medarbBelob 4 = Real. omsætning lukkede job
    '** medarbBelob 5 = Real. internKost lukkede job
    '** medarbBelob 6 = Faktureret lukkede job

    '** medarbBelob 9 = Real. timer
    '** medarbBelob 10 = Salgsomkostninger lukkede job
    '** medarbBelob 11 = Faktureret, exlc. salgsomkostninger og km. mv. lukkede job


    '** medarbBelob 12 = Real. timer alle åbne job
  

    '** medarbBelob 7 = InternKost alle åbne job
    '** medarbBelob 8 = Faktureret alle åbne job

    '** medarbBelob 13 = Salgsomkostninger åbne job
    

    '** medarbBelob 14 = til Effektiv tp 
    '** medarbBelob 15 = tot timer pr. medarb i periode til beregning af faktor 


    'strExportA(0,0) = "Medarb./Kategorier"
    
    lastMid = 0
    'isJbWrt = "##0##"
    thisBel = 0
    call akttyper2009(2)

    strSQLm = "SELECT mid FROM medarbejdere AS m WHERE "& medarbSQlKri &") ORDER BY mnavn"
    
    'Response.write strSQLm
    'Response.flush

    mids = 0
  
    oRec2.open strSQLm, oConn, 3
    while not oRec2.EOF

    jobisWrt = "##0##"
    m = 1
    meStamdata(oRec2("mid"))
    strSQLjob = ""
    salgsOmkostIaltprMedarb = 0
    kostIaltprMedarb = 0
    omsIaltprMedarb = 0
    realTimerIaltprMedarb = 0
    useJobSQL = " AND (j.id = 0 "
    faktor = 0

    'medarbTimer = 0
    'fakkunTimer = 0
    'matkobspris = 0
    

    if media <> "export" AND media <> "chart" then%>
    <tr bgcolor="#FFFFFF">
    <td style="white-space:nowrap;"><%=meTxt%> </td>
    <%end if %>

    <%
    
   ' call jq_format(meTxt)
    'meTxt = jq_formatTxt
    strExport = strExport & meTxt &";" 

    strExportA(mids,0) = meTxt
   
    'Grundlag = ""
    %>

                            <%for m = 0 to 10 
                              formelSQL = ""
                             'thisBel = 0
                             'medarbBelob(oRec2("mid"),m) = 0

                            

                                                         select case m
                                                         

                                                         case 0
                                                         if cint(viskol1) = 1 then
                                                         '** tilbud unden sandsynkligheds beregning ***
                                                         if toplinie_db = 0 then 'db

                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100), 0), '0') AS td1_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100), 0), '0') AS td2_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100), 0), '0') AS td3_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100), 0), '0') AS td4_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100), 0), '0') AS td5_oms"

                                                         else
                                                         
                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100), 0), '0') AS td1_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100), 0), '0') AS td2_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100), 0), '0') AS td3_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100), 0), '0') AS td4_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100), 0), '0') AS td5_oms"

                                                         end if

                                                         jobStKriSQL = "j_td1.jobstartdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND j_td1.jobstatus = 3 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         fakSQL = ""
                                                         joinValutaFlt = "j_td1.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""
                                                         
                                                         call salgogvaerdiSQL

                                                         end if
                                                         
                                                         
                                                         
                                                         
                                                         case 1
                                                         
                                                         
                                                         if cint(viskol1) = 1 then   
                                                         if toplinie_db = 0 then
                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td1_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td2_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td3_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td4_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         & "ROUND((((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td5_oms"
                                                         
                                                         else
                                                         
                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td1_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td2_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td3_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td4_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         & "ROUND(((j_td1.jo_bruttooms * v.kurs)/100) * (j_td1.sandsynlighed / 100), 0), '0') AS td5_oms"

                                                         
                                                         end if

                                                         jobStKriSQL = "j_td1.jobstartdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND j_td1.jobstatus = 3 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         fakSQL = ""
                                                         joinValutaFlt = "j_td1.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""

                                                         call salgogvaerdiSQL
                                                         

                                                         end if


                                                         
                                                         case 2


                                                         if cint(viskol1) = 1 then
                                                         formelSQL = "SUM(t.timer) AS realtimer"
     
                                                         if toplinie_db = 0 then 'db
                                                         
                                                        

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         & " if (stade_tim_proc = 1, ROUND((restestimat/100) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td1_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         & " if (stade_tim_proc = 1, ROUND((restestimat/100) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td2_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         & " if (stade_tim_proc = 1, ROUND((restestimat/100) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td3_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         & " if (stade_tim_proc = 1, ROUND((restestimat/100) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td4_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         & " if (stade_tim_proc = 1, ROUND((restestimat/100) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td5_oms"
                                                         

                                                         else
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" if (stade_tim_proc = 1, ROUND((restestimat/100) * ((j_td1.jo_bruttooms * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td1_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" if (stade_tim_proc = 1, ROUND((restestimat/100) * ((j_td1.jo_bruttooms * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td2_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" if (stade_tim_proc = 1, ROUND((restestimat/100) * ((j_td1.jo_bruttooms * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td3_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" if (stade_tim_proc = 1, ROUND((restestimat/100) * ((j_td1.jo_bruttooms * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td4_oms"
                                                         

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" if (stade_tim_proc = 1, ROUND((restestimat/100) * ((j_td1.jo_bruttooms * v.kurs)/100),0), "_
                                                         &" ROUND((SUM(t.timer)/(restestimat+SUM(t.timer))) * (((j_td1.jo_bruttooms - (j_td1.jo_udgifter_intern + j_td1.jo_udgifter_ulev)) * v.kurs)/100),0)), '0') AS td5_oms"
                                                         

                                                         end if

                                                         jobStKriSQL = "j_td1.jobstartdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND j_td1.jobstatus = 1 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         fakSQL = " LEFT JOIN timer AS t ON (t.tjobnr = j_td1.jobnr AND ("& aty_sql_realhours &"))"
                                                         joinValutaFlt = "j_td1.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""
                                                         
                                                         call salgogvaerdiSQL

                                                         end if

                                                         case 3

                                                          'if cint(viskol5) = 1 then

                                                         '** Salgs omk lukkede job ***
                                                         formelSQL = formelSQL & " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_1/100),0),0) AS matkobspris_1"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_2/100),0),0) AS matkobspris_2"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_3/100),0),0) AS matkobspris_3"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_4/100),0),0) AS matkobspris_4"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_5/100),0),0) AS matkobspris_5"

                                                         jobStKriSQL = "j_td1.jobslutdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND j_td1.jobstatus = 0 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         
                                                         fakSQL = " LEFT JOIN materiale_forbrug AS mf ON (mf.jobid = j_td1.id)"
                                                         joinValutaFlt = "mf.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""


                                                         call salgogvaerdiSQL


                                                         
                                                         case 4 '***** salgs timepris (real. oms) lukkede job i periode 

                                                         if cint(viskol2) = 1 then

                                                         '*** Real timer
                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_1/100),0),0) AS realtimer_1"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_2/100),0),0) AS realtimer_2"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_3/100),0),0) AS realtimer_3"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_4/100),0),0) AS realtimer_4"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_5/100),0),0) AS realtimer_5"


                                                         '*** Omsætning
                                                         formelSQL = formelSQL &  ", if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_1/100),0),0) AS td1_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_2/100),0),0) AS td2_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_3/100),0),0) AS td3_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_4/100),0),0) AS td4_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_5/100),0),0) AS td5_oms"

                                                        

                                                         '*** Kostpriser
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_1/100),0),0) AS td1_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_2/100),0),0) AS td2_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_3/100),0),0) AS td3_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_4/100),0),0) AS td4_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_5/100),0),0) AS td5_kost"


                                                         jobStKriSQL = "j_td1.jobslutdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND j_td1.jobstatus = 0 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         fakSQL = " LEFT JOIN timer AS t ON (t.tjobnr = j_td1.jobnr AND ("& aty_sql_realhours &"))"
                                                         'fakSQL = fakSQL & " LEFT JOIN materiale_forbrug AS mf ON (mf.jobid = j_td1.id)"
                                                         joinValutaFlt = "j_td1.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""


                                                         call salgogvaerdiSQL

                                                         end if

                                                         case 5 '*** Faktureret oms lukkede job

                                                         if cint(viskol2) = 1 then
                                                       

                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_1/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_1/100)),0)), '0') AS td1_oms"
                                                        
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_2/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_2/100)),0)), '0') AS td2_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_3/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_3/100)),0)), '0') AS td3_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_4/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_4/100)),0)), '0') AS td4_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_5/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_5/100)),0)), '0') AS td5_oms"
                                                        

                                                         
                                                      
                                                         

                                                         jobStKriSQL = "j_td1.jobstatus = 0 AND j_td1.risiko > -1 AND jobslutdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND (" & strKnrSQLkri & ")"
                                                         fakSQL = "LEFT JOIN fakturaer AS f ON ((f.jobid = j_td1.id OR (f.aftaleid = j_td1.serviceaft AND f.aftaleid <> 0)) AND (f.medregnikkeioms <> 1 AND f.medregnikkeioms <> 2 AND f.shadowcopy = 0)) "
                                                         fakSQL = fakSQL &" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"

                                                         
                                                       
                                                         joinValutaFlt = "f.valuta"
                                                         sqlGrpBy = "j_td1.id, fid"
                                                         whEkstra = " AND ((f.jobid = j_td1.id OR (f.aftaleid = j_td1.serviceaft AND f.aftaleid <> 0)))"
                                                         

                                                         call salgogvaerdiSQL

                                                         end if

                                                         case 6 '**** Internkost, Realiseret oms, timer Åbne job

                                                         if cint(viskol3) = 1 then

                                                           '*** Real timer
                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_1/100),0),0) AS realtimer_1"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_2/100),0),0) AS realtimer_2"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_3/100),0),0) AS realtimer_3"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_4/100),0),0) AS realtimer_4"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM(t.timer * j_td1.jobans_proc_5/100),0),0) AS realtimer_5"

                                                         '*** Omsætning
                                                         formelSQL = formelSQL &  ", if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_1/100),0) AS td1_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_2/100),0) AS td2_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_3/100),0) AS td3_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_4/100),0) AS td4_oms"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.timepris * (t.kurs/100)) * (j_td1.jobans_proc_5/100),0) AS td5_oms"

                                                         '** Salgs omk ***
                                                         'formelSQL = formelSQL & ", ROUND(SUM((matkobspris*matantal) * v.kurs/100),0) AS matkobspris "

                                                         '*** Kostpriser
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_1/100),0) AS td1_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_2/100),0) AS td2_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_3/100),0) AS td3_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_4/100),0) AS td4_kost"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" SUM(t.timer * t.kostpris * (t.kurs/100)) * (j_td1.jobans_proc_5/100),0) AS td5_kost"


                                                         jobStKriSQL = "j_td1.jobslutdato >= '"& stDato &"' AND j_td1.jobstatus = 1 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         fakSQL = " LEFT JOIN timer AS t ON (t.tjobnr = j_td1.jobnr AND ("& aty_sql_realhours &"))"
                                                         'fakSQL = fakSQL & " LEFT JOIN materiale_forbrug AS mf ON (mf.jobid = j_td1.id)"
                                                         joinValutaFlt = "j_td1.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""

                                                         call salgogvaerdiSQL

                                                         end if

                                                         case 7 '**** Faktureret omsætning Åbne job                                                         

                                                         if cint(viskol3) = 1 then

                                                       
                                                       

                                                         formelSQL = " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_1/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_1/100)),0)), '0') AS td1_oms"
                                                        
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_2/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_2/100)),0)), '0') AS td2_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_3/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_3/100)),0)), '0') AS td3_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_4/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_4/100)),0)), '0') AS td4_oms"
                                                         
                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" if (f.faktype <> 1, ROUND(SUM((f.beloeb * f.kurs)/100*(j_td1.jobans_proc_5/100)),0), ROUND(SUM((f.beloeb * -1 * f.kurs)/100 * (j_td1.jobans_proc_5/100)),0)), '0') AS td5_oms"
                                                        

                                                         
                                                      
                                                         

                                                         jobStKriSQL = "j_td1.jobstatus = 1 AND j_td1.risiko > -1 AND j_td1.jobslutdato >= '"& stDato &"' AND (" & strKnrSQLkri & ")"
                                                         fakSQL = "LEFT JOIN fakturaer AS f ON ((f.jobid = j_td1.id OR (f.aftaleid = j_td1.serviceaft AND f.aftaleid <> 0)) AND (f.medregnikkeioms = 0 AND f.shadowcopy = 0)) "
                                                         fakSQL = fakSQL &" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"

                                                         
                                                         joinValutaFlt = "f.valuta"
                                                         sqlGrpBy = "j_td1.id, fid"
                                                         whEkstra = " AND ((f.jobid = j_td1.id OR (f.aftaleid = j_td1.serviceaft AND f.aftaleid <> 0)))"

                                                         call salgogvaerdiSQL

                                                         end if

                                                         case 8

                                                         
                                                         'end if

                                                          case 9

                                                          if cint(viskol5) = 1 then

                                                         '** Salgs omk åbne job ***
                                                         formelSQL = formelSQL & " if ( j_td1.jobans1 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_1/100),0),0) AS matkobspris_1"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans2 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_2/100),0),0) AS matkobspris_2"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans3 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_3/100),0),0) AS matkobspris_3"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans4 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_4/100),0),0) AS matkobspris_4"

                                                         formelSQL = formelSQL & ", if ( j_td1.jobans5 = "& oRec2("mid") &", "_
                                                         &" ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)) * (j_td1.jobans_proc_5/100),0),0) AS matkobspris_5"

                                                         jobStKriSQL = "j_td1.jobslutdato BETWEEN '"& stDato &"' AND '"& slDato &"' AND j_td1.jobstatus = 1 AND j_td1.risiko > -1 AND (" & strKnrSQLkri & ")"
                                                         
                                                         fakSQL = " LEFT JOIN materiale_forbrug AS mf ON (mf.jobid = j_td1.id)"
                                                         joinValutaFlt = "mf.valuta"
                                                         sqlGrpBy = "j_td1.id"
                                                         whEkstra = ""


                                                         call salgogvaerdiSQL

                                                         end if

                                                         case 10

                                                            if cint(viskol6) = 1 then

                                                            '*** Henter alle job med timer på, der er lukket siden stratdato i interval ***'
                                                            strSQLjobA = "SELECT j.id AS jid, ROUND(SUM(t.timer), 0) AS realtimerprmedarb "_
                                                            &" FROM job AS j "_
                                                            &" LEFT JOIN timer AS t ON (t.tjobnr = j.jobnr AND ("& aty_sql_realHoursFakbar &") AND t.tmnr = " & oRec2("mid") & ") "_
                                                            &" WHERE j.jobstatus = 0 AND j.risiko > -1 AND j.jobslutdato >= '"& stDato &"' AND t.tjobnr = j.jobnr AND t.timer > 0 GROUP BY j.id ORDER BY j.id"

                                                           

                                                            oRec3.open strSQLjobA, oConn, 3

                                                            while not oRec3.EOF

                                                            useJobSQL = useJobSQL & " OR j.id = " & oRec3("jid")
                                                            jobidsEgneTimer(oRec3("jid")) = oRec3("realtimerprmedarb")
                                                            
                                                            oRec3.movenext
                                                            wend
                                                            oRec3.close
                                                            
                                                            useJobSQL = useJobSQL & ")"


                                                            '** Henter Mat forbrug ***'
                                                            strSQLjobB = "SELECT jobid, ROUND(SUM((matkobspris*matantal) * (mf.kurs/100)),0) AS salgsomk"_
                                                            &" FROM materiale_forbrug AS mf WHERE id <> 0 AND jobid <> 0  "& replace(useJobSQL, "j.id", "mf.jobid") & " GROUP BY mf.jobid"
                                                            
                                                            'Response.Write "<br><br>"& strSQLjobB  & "<br><br>"
                                                            'Response.end
                                                            oRec3.open strSQLjobB, oConn, 3

                                                            while not oRec3.EOF

                                                            jobidsSalgsomkost(oRec3("jobid")) = oRec3("salgsomk")
                                                            
                                                            oRec3.movenext
                                                            wend
                                                            oRec3.close            


                                                             '** Henter Fakturaer ***'
                                                            strSQLjobC = "SELECT jobid, if (f.faktype <> 1, ROUND((f.beloeb * f.kurs)/100,0), ROUND((f.beloeb * -1 * f.kurs)/100,0)) AS faktureret"_
                                                            &" FROM fakturaer AS f WHERE fid <> 0 AND jobid <> 0  "& replace(useJobSQL, "j.id", "f.jobid") & " ORDER BY f.jobid"
                                                            
                                                            'Response.Write "<br><br>"& strSQLjobC  & "<br><br>"
                                                            'Response.end
                                                            oRec3.open strSQLjobC, oConn, 3

                                                            while not oRec3.EOF

                                                            jobidsFaktureret(oRec3("jobid")) = jobidsFaktureret(oRec3("jobid")) + (oRec3("faktureret"))
                                                            
                                                            oRec3.movenext
                                                            wend
                                                            oRec3.close      

                                                            strSQLjob = "SELECT kkundenavn, kkundenr, jobnavn, j.id AS jid, jobnr, ROUND(SUM(t2.timer * t2.kostpris * (t2.kurs/100)),0) AS realomk, j.virksomheds_proc AS td1_virksomhedsproc "_
                                                            &" FROM job AS j "_
                                                            &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                                                            &" LEFT JOIN timer AS t2 ON (t2.tjobnr = j.jobnr AND ("& replace(aty_sql_realHoursFakbar, "t.", "t2.") &")) "_
                                                            &" WHERE jobstatus = 0 "&useJobSQL&" AND t2.timer > 0 GROUP BY j.id ORDER BY j.id"

                                                          



                                                            end if

                                                         end select






                                                       

                                                        'if lto = "epi" then
                                                        'Response.write "foreml: " & formelSQL & "<br>"
                                                        'if oRec2("mid") = 1 then
                                                        'if m = 8 then
                                                        'Response.write "<br>"& m &"<br>"& strSQLjob & "<br><br>"
                                                        'Response.flush
                                                        'end if
                                                        
                                                        'end if
                                                        'end if
    
                                                        'matkobspris = 0

                                                         
                                                         if strSQLjob <> "" then
                                                                    
                                                                    j = 0
                                                                    procIalt = 0 '*Tjeker at en medab ikke overstiger 100% på et job
                                                                    oRec.open strSQLjob, oConn, 3
                                                                    while not oRec.EOF

                                                                    grThisbelob = ""
                                                                    
                                                                        
                                                                        'Response.write "hej<br>"


                                                                        if m <> 3 AND m <> 8 AND m <> 9 AND m <> 10 then
                                                                        '**** Real omsætning pr. medarb
                                                                        if len(trim(oRec("td1_oms"))) then
                                                                        cdblOms1 = formatnumber(oRec("td1_oms"), 0)
                                                                        else
                                                                        cdblOms1 = 0
                                                                        end if

                                                                         if len(trim(oRec("td2_oms"))) then
                                                                        cdblOms2 = formatnumber(oRec("td2_oms"), 0)
                                                                        else
                                                                        cdblOms2 = 0
                                                                        end if


                                                                        if len(trim(oRec("td3_oms"))) then
                                                                        cdblOms3 = formatnumber(oRec("td3_oms"), 0)
                                                                        else
                                                                        cdblOms3 = 0
                                                                        end if


                                                                         if len(trim(oRec("td4_oms"))) then
                                                                        cdblOms4 = formatnumber(oRec("td4_oms"), 0)
                                                                        else
                                                                        cdblOms4 = 0
                                                                        end if

                                                                         if len(trim(oRec("td5_oms"))) then
                                                                        cdblOms5 = formatnumber(oRec("td5_oms"), 0)
                                                                        else
                                                                        cdblOms5 = 0
                                                                        end if

                                                                        omsIaltprMedarb = formatnumber(cdblOms1/1 + cdblOms2/1 + cdblOms3/1 + cdblOms4/1 + cdblOms5/1, 0) 
                                                                        end if


                                                                        select case m
                                                                        case 4,6
                                                                        
                                                                        '***** Real kostpriser pr medarb
                                                                        if len(trim(oRec("td1_kost"))) then
                                                                        cdblKost1 = formatnumber(oRec("td1_kost"), 0)
                                                                        else
                                                                        cdblKost1 = 0
                                                                        end if

                                                                         if len(trim(oRec("td2_kost"))) then
                                                                        cdblKost2 = formatnumber(oRec("td2_kost"), 0)
                                                                        else
                                                                        cdblKost2 = 0
                                                                        end if


                                                                        if len(trim(oRec("td3_kost"))) then
                                                                        cdblKost3 = formatnumber(oRec("td3_kost"), 0)
                                                                        else
                                                                        cdblKost3 = 0
                                                                        end if


                                                                         if len(trim(oRec("td4_kost"))) then
                                                                        cdblKost4 = formatnumber(oRec("td4_kost"), 0)
                                                                        else
                                                                        cdblKost4 = 0
                                                                        end if

                                                                         if len(trim(oRec("td5_kost"))) then
                                                                        cdblKost5 = formatnumber(oRec("td5_kost"), 0)
                                                                        else
                                                                        cdblKost5 = 0
                                                                        end if

                                                                        ejerPrcentIalt = (oRec("td1_proc1") + oRec("td1_proc2") + oRec("td1_proc3") + oRec("td1_proc4") + oRec("td1_proc5"))/1
                                                                        kostIaltprMedarb = formatnumber(cdblKost1/1 + cdblKost2/1 + cdblKost3/1 + cdblKost4/1 + cdblKost5/1, 0) 

                                                                       
                                                                        '***** Real timer pr medarb
                                                                        if len(trim(oRec("realtimer_1"))) then
                                                                        cdblRealTimer1 = formatnumber(oRec("realtimer_1"), 0)
                                                                        else
                                                                        cdblRealTimer1 = 0
                                                                        end if

                                                                         if len(trim(oRec("realtimer_2"))) then
                                                                        cdblRealTimer2 = formatnumber(oRec("realtimer_2"), 0)
                                                                        else
                                                                        cdblRealTimer2 = 0
                                                                        end if


                                                                        if len(trim(oRec("realtimer_3"))) then
                                                                        cdblRealTimer3 = formatnumber(oRec("realtimer_3"), 0)
                                                                        else
                                                                        cdblRealTimer3 = 0
                                                                        end if


                                                                        if len(trim(oRec("realtimer_4"))) then
                                                                        cdblRealTimer4 = formatnumber(oRec("realtimer_4"), 0)
                                                                        else
                                                                        cdblRealTimer4 = 0
                                                                        end if

                                                                         if len(trim(oRec("realtimer_5"))) then
                                                                        cdblRealTimer5 = formatnumber(oRec("realtimer_5"), 0)
                                                                        else
                                                                        cdblRealTimer5 = 0
                                                                        end if

                                                                        realTimerIaltprMedarb = formatnumber(cdblRealTimer1/1 + cdblRealTimer2/1 + cdblRealTimer3/1 + cdblRealTimer4/1 + cdblRealTimer5/1, 0) 

                                                                       

                                                                        'salgsOmkostIaltprMedar = 0

                                                                        case 3, 9

                                                                        
                                                                        '***** Salgsomkostninger pr medarb
                                                                        if len(trim(oRec("matkobspris_1"))) <> 0 then
                                                                        cdblmat1 = formatnumber(oRec("matkobspris_1"), 0)
                                                                        else
                                                                        cdblmat1 = 0
                                                                        end if

                                                                         if len(trim(oRec("matkobspris_2"))) <> 0 then
                                                                        cdblmat2 = formatnumber(oRec("matkobspris_2"), 0)
                                                                        else
                                                                        cdblmat2 = 0
                                                                        end if


                                                                        if len(trim(oRec("matkobspris_3"))) <> 0 then
                                                                        cdblmat3 = formatnumber(oRec("matkobspris_3"), 0)
                                                                        else
                                                                        cdblmat3 = 0
                                                                        end if


                                                                        if len(trim(oRec("matkobspris_4"))) <> 0 then
                                                                        cdblmat4 = formatnumber(oRec("matkobspris_4"), 0)
                                                                        else
                                                                        cdblmat4 = 0
                                                                        end if

                                                                        if len(trim(oRec("matkobspris_5"))) <> 0 then
                                                                        cdblmat5 = formatnumber(oRec("matkobspris_5"), 0)
                                                                        else
                                                                        cdblmat5 = 0
                                                                        end if

                                                                        'Response.Write "her:" & oRec("matkobspris_1") & "<br>"

                                                                        salgsOmkostIaltprMedarb = formatnumber(cdblmat1/1 + cdblmat2/1 + cdblmat3/1 + cdblmat4/1 + cdblmat5/1, 0) 


                                                                        case 10

                                                                        faktor = 0
                                                                        
                                                                        'Response.write oRec("jobnavn") & " Fak:"& jobidsFaktureret(oRec("jid")) &" / Salgsomk: "& jobidsSalgsomkost(oRec("jid")) &" + Intern omk: "& oRec("realomk") &" * Egnetimer: "& jobidsEgneTimer(oRec("jid")) & "<br>"

                                                                        if (jobidsSalgsomkost(oRec("jid")) + oRec("realomk")) <> 0 then
                                                                        faktor = ((jobidsFaktureret(oRec("jid")) - (jobidsSalgsomkost(oRec("jid")))) /  oRec("realomk")) * jobidsEgneTimer(oRec("jid")) 'oRec("realtimerprmedarb")
                                                                        else
                                                                        faktor = 1
                                                                        end if

                                                                        'Response.Write "faktor f. timer:"& formatnumber((jobidsFaktureret(oRec("jid")) / (jobidsSalgsomkost(oRec("jid")) + oRec("realomk"))), 2) & " faktor:" & faktor  &"  Egnetimer: "& jobidsEgneTimer(oRec("jid")) & "<br>"
                                                                        realTimerIaltprMedarb = jobidsEgneTimer(oRec("jid")) 'oRec("realtimerprmedarb") 

                                                                        end select


        
        
                                                                                   
                                                                                   
                                                                                    if cint(virkAndel) = 1 then '** benyt virksomhedsandel

                                                                                        '**Hvis andel = 100%  ==> 0% til medarbejer
                                                                                        if oRec("td1_virksomhedsproc") <> 0 then
                         
                                                                                             virksomhedsproc = 1 - (oRec("td1_virksomhedsproc")/100)
                      
                                                                                        else
                                                                                        virksomhedsproc = 1
                                                                                        end if

                                                                                    else

                                                                                    virksomhedsproc = 1

                                                                                    end if






                                                                                    select case m
                                                                                   

                                                                                    case 0
                                                                                        
                                                                                        if cint(viskol1) = 1 then

                                                                                        if grTxt1_fundet <> 1 then
                                                                                        grTxt = ""'<br><br>Tilbud: "
                                                                                        grTxt1_fundet = 1
                                                                                        else
                                                                                        grTxt = ""
                                                                                        end if

                                                                                        grThisbelob = formatnumber((omsIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO

                                                                                        end if

                                                                                    case 2
                                                                                       
                                                                                       if cint(viskol1) = 1 then

                                                                                       if grTxt3_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Job (WIP): "
                                                                                       grTxt3_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if
                                                                                       
                                                                                       grThisbelob = formatnumber((omsIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO
                                                                                       
                                                                                       end if

                                                                                    
                                                                                     case 3
                                                                                        
                                                                                          'if cint(viskol5) = 1 then
                                                                                       'Response.write "herder69<br>"
                                                                                        medarbBelob(oRec2("mid"),10) = medarbBelob(oRec2("mid"),10) + (salgsOmkostIaltprMedarb/1 * virksomhedsproc/1)


                                                                                        if grTxt8_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Salgsomkostninger lukkede job: "
                                                                                       grTxt8_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if
                                                                                      

                                                                                       grThisbelob = formatnumber((salgsOmkostIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO

                                                                                       'end if


                                                                                    case 4

                                                                                    if cint(viskol2) = 1 then
                                                                                    medarbBelob(oRec2("mid"),9) = medarbBelob(oRec2("mid"),9)/1 + (realTimerIaltprMedarb/1 * virksomhedsproc/1)


                                                                                    'Response.Write medarbBelob(oRec2("mid"),9) & " Real: " & oRec("realtimer") & "<br>"

                                                                                    

                                                                                    medarbBelob(oRec2("mid"),5) = medarbBelob(oRec2("mid"),5) + (kostIaltprMedarb/1 * virksomhedsproc/1) 
                                                                                    medarbBelob(oRec2("mid"),4) = medarbBelob(oRec2("mid"),4) + (omsIaltprMedarb/1 * virksomhedsproc/1)

                                                                                   

                                                                                       if grTxt4_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Real. lukkede job: "
                                                                                       grTxt4_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if

                                                                                        if realTimerIaltprMedarb <> 0 then
                                                                                        grThisbelob = "Timer: "& formatnumber(realTimerIaltprMedarb/1 * virksomhedsproc/1,0) & ", Real. Oms.: "& formatnumber((omsIaltprMedarb/1 * virksomhedsproc/1),0) & " "& basisValISO &", Real. Omkost.: "& formatnumber((kostIaltprMedarb/1 * virksomhedsproc/1),0) & " DKK, Andel %: "& formatnumber(ejerPrcentIalt,0) &"<br>"
                                                                                        else
                                                                                        grThisbelob = ""
                                                                                        end if

                                                                                        end if

                                                                                    case 5

                                                                                        if cint(viskol2) = 1 then

                                                                                        medarbBelob(oRec2("mid"),6) = medarbBelob(oRec2("mid"),6) + (omsIaltprMedarb/1 * virksomhedsproc/1)
                                                                                        
                                                                                        

                                                                                        select case lto
                                                                                        case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab"
                                                                                        medarbBelob(oRec2("mid"),11) = medarbBelob(oRec2("mid"),11) + (omsIaltprMedarb/1 * virksomhedsproc/1) '+ medarbBelob(oRec2("mid"),6) '' * (virksomhedsproc/1)) 'medarbBelob(oRec2("mid"),m)
                                                                                        case else
                                                                                        medarbBelob(oRec2("mid"),11) = medarbBelob(oRec2("mid"),11) + (oRec("aktbel") * virksomhedsproc/1)
                                                                                        end select
                                                                                    
                                                                                    
                                                                                        
                                                                                       if grTxt5_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Faktureret, lukkede job: "
                                                                                       grTxt5_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if

                                                                                        if omsIaltprMedarb <> 0 then
                                                                                        grThisbelob = "Faktureret: "& formatnumber((omsIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO
                                                                                        else
                                                                                        grThisbelob = ""
                                                                                        end if

                                                                                        end if

                                                                                    case 6

                                                                                    if cint(viskol3) = 1 then

                                                                                    medarbBelob(oRec2("mid"),7) = medarbBelob(oRec2("mid"),7) + (kostIaltprMedarb/1 * virksomhedsproc/1) 
                                                                                    medarbBelob(oRec2("mid"),12) = medarbBelob(oRec2("mid"),12) + (realTimerIaltprMedarb/1 * virksomhedsproc/1)

                                                                                    
                                                                                        
                                                                                       if grTxt6_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Real. åbne job: "
                                                                                       grTxt6_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if

                                                                                        'grThisbelob = formatnumber((kostIaltprMedarb/1 * virksomhedsproc/1),0) & " DKK"
                                                                                        if realTimerIaltprMedarb <> 0 then
                                                                                        grThisbelob = "Timer: "& formatnumber(realTimerIaltprMedarb/1 * virksomhedsproc/1,0) & ", Real. Oms.: "& formatnumber((omsIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO &", Real. Omkost.: "& formatnumber((kostIaltprMedarb/1 * virksomhedsproc/1),0) & " DKK, Andel %: "& formatnumber(ejerPrcentIalt,0) &"<br>"
                                                                                        else
                                                                                         grThisbelob = ""
                                                                                        end if

                                                                                    end if
                                                                                    case 7

                                                                                    if cint(viskol3) = 1 then

                                                                                           if grTxt7_fundet <> 1 then
                                                                                           grTxt = ""'<br><br>Faktureret, åbne job: "
                                                                                           grTxt7_fundet = 1
                                                                                           else
                                                                                           grTxt = ""
                                                                                           end if

                                                                                          medarbBelob(oRec2("mid"),8) = medarbBelob(oRec2("mid"),8) + (omsIaltprMedarb/1 * virksomhedsproc/1)

                                                                                          if omsIaltprMedarb <> 0 then
                                                                                          grThisbelob = "Faktureret: "& formatnumber((omsIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO
                                                                                          else
                                                                                          grThisbelob = ""
                                                                                          end if

                                                                                       end if
                                                                                    
                                                                                     case 8
                                                                                        
                                                                                       
                                                                                     

                                                                                     case 9
                                                                                        

                                                                                          if cint(viskol6) = 1 then
                                                                                       'Response.write "herder69<br>"
                                                                                        medarbBelob(oRec2("mid"),13) = medarbBelob(oRec2("mid"),13) + (salgsOmkostIaltprMedarb/1 * virksomhedsproc/1)


                                                                                       if grTxt9_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Salgsomkostninger åbne job: "
                                                                                       grTxt9_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if
                                                                                       
                                                                                        grThisbelob = formatnumber((salgsOmkostIaltprMedarb/1 * virksomhedsproc/1),0) &" "& basisValISO
                                                                                        
                                                                                        end if

                                                                                    
                                                                                    case 10
                                                                                      
                                                                                        if cint(viskol6) = 1 then

                                                                                      '** Faktor / Virksomheds % medregnes ikke her ***'
                                                                                      medarbBelob(oRec2("mid"),14) = medarbBelob(oRec2("mid"),14) + (faktor/1)
                                                                                      medarbBelob(oRec2("mid"),15) = medarbBelob(oRec2("mid"),15) + (realTimerIaltprMedarb/1)

                                                                                       if grTxt10_fundet <> 1 then
                                                                                       grTxt = ""'<br><br>Faktor ALLE job: "
                                                                                       grTxt10_fundet = 1
                                                                                       else
                                                                                       grTxt = ""
                                                                                       end if
                                                                                       
                                                                                       if faktor <> 0 then
                                                                                       faktor = formatnumber(faktor,0)
                                                                                       else
                                                                                       faktor = 0
                                                                                       end if

                                                                                       if realTimerIaltprMedarb <> 0 then
                                                                                       realTimerIaltprMedarb = formatnumber(realTimerIaltprMedarb, 0) 
                                                                                       'grThisbelob = realTimerIaltprMedarb & " t." 'formatnumber(faktor, 0) &"/"& /realTimerIaltprMedarb
                                                                                      
                                                                                       'if instr(jobisWrt, "#"&oRec("jid")&"#") = 0 then
                                                                                       'jobisWrt = jobisWrt & ",#"&oRec("jid")&"#"
                                                                                       grThisbelob = "Faktor: " & formatnumber((jobidsFaktureret(oRec("jid")) / (jobidsSalgsomkost(oRec("jid")) + oRec("realomk"))), 2)  & " Egne timer: "& jobidsEgneTimer(oRec("jid")) &", Faktureret: "& formatnumber(jobidsFaktureret(oRec("jid")),0) &", Salgsomk.: "& formatnumber(jobidsSalgsomkost(oRec("jid")),0) &", Real. Omk.: "& formatnumber(oRec("realomk"),0) &"<br>" 
                                                                                       ' <br>: Faktor 14: " & formatnumber(medarbBelob(oRec2("mid"),14), 2) & " egne timer: "& medarbBelob(oRec2("mid"),15) & " egne timer ialt:" & grThisbelob & "<br>" 
                                                                                        
                                                                                       'else
                                                                                       'grThisbelob = ""
                                                                                       'end if

                                                                                       else
                                                                                       realTimerIaltprMedarb = 0
                                                                                       grThisbelob = 0
                                                                                       end if

                                                                                        

                                                                                        end if

                                                                                    end select
                                                                                    
                                                                                   
                                                                                    if grThisbelob <> "0 "& basisValISO AND len(trim(grThisbelob)) <> 0 then 'cint(viskol4) = 1 AND 
                                                                                    if media <> "export" AND media <> "chart" then
                                                                                    medarbGrlag(oRec2("mid"),m) = medarbGrlag(oRec2("mid"),m) &"<u>"& grTxt &"</u><br> <b>"& left(oRec("kkundenavn"), 20)&" | "& left(oRec("jobnavn"), 15) & "</b> ("& oRec("jobnr") &")<br>"& grThisbelob   
                                                                                    end if
                                                                                    end if


                                                                       
    
                                                                    j = j + 1
                                                                    lastMid = oRec2("mid")   
                                                                    procIalt = 0 
                                                                    oRec.movenext
                                                                    wend 
                                                                    oRec.close

                                                                    end if
                                                                    strSQLjob = ""
    


                                

                                'thisBelob = thisbelob + medarbBelob(oRec2("mid"),m)
                                'Response.Write "<br><b>"& medarbBelob(oRec2("mid"),m) & " "& oRec2("mid") &"</b> xDKK<br>"

                          
                                
                            
                                
                                   
                                    

                                                      



                                                '******* Kolonne værdier ****'

                                                select case m
                                               
                                                case 0
                                                kol1_belob = (medarbBelob(oRec2("mid"),m)/1) 
                                                case 1
                                                kol2_belob = (medarbBelob(oRec2("mid"),m)/1) 
                                                case 2
                                                kol3_belob = (medarbBelob(oRec2("mid"),m)/1) 

                                                case 3
                                                 '*** Salgsomkostninger lukkede job
                                                kol10_belob = (medarbBelob(oRec2("mid"),10)/1) 
                                                case 4 '*** Lukkede job

                                                    kol4_belob = (medarbBelob(oRec2("mid"),m)/1) 

                                                    '*** Real. timer
                                                    kol9_belob = (medarbBelob(oRec2("mid"),9)/1)
                                                   
                                                    
                                                    '*** Kost lukkede job i periode
                                                    kol5_belob = (medarbBelob(oRec2("mid"),5)/1) 

                                                
                                                case 5
                                                

                                                '**** Faktureret kun timer ( lukekde job i periode)
                                                kol11_belob = (medarbBelob(oRec2("mid"),11)/1)
                                                
                                                '**** Faktureret oms (lukkede job i periode)
                                                kol6_belob = (medarbBelob(oRec2("mid"),6)/1) 

                                                

                                                case 6
                                                kol7_belob = (medarbBelob(oRec2("mid"),7)/1) 
                                                kol12_belob = (medarbBelob(oRec2("mid"),12)/1) 

                                                case 7
                                                '***** Faktureret oms (Åbne job i periode)
                                                kol8_belob = (medarbBelob(oRec2("mid"),8)/1) 
                                                
                                                case 8
                                               
                                                
                                                case 9  
                                                
                                                '*** Salgsomkostninger åbne job
                                                kol13_belob = (medarbBelob(oRec2("mid"),13)/1) 

                                                case 10
                                                
                                                if medarbBelob(oRec2("mid"),15) <> 0 then
                                                kol14_belob = (medarbBelob(oRec2("mid"),14)/medarbBelob(oRec2("mid"),15)) 
                                                else
                                                kol14_belob = 0
                                                end if
                                                   
                                                end select


                                        
                                                             'if medarbBelob(oRec2("mid"),m) <> 0 then
                                                 
                                                 if media <> "export" AND media <> "chart" then   
                                                     
                                                             select case m
                                                             case 0
                                                             
                                                             if cint(viskol1) = 1 then
                                                             %>
                                                              <td align=right style="white-space:nowrap;">
                                                              <%if kol1_belob <> 0 then %>
                                                             <%=formatnumber(kol1_belob,0) %>
                                                             <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%
                                                             kol1_belobTot = kol1_belobTot + kol1_belob  
                                                             end if

                                                              case 1
                                                             
                                                             if cint(viskol1) = 1 then

                                                              %>
                                                             
                                                              <td align=right style="white-space:nowrap;">
                                                              <%if kol2_belob <> 0 then %>
                                                             <%=formatnumber(kol2_belob,0) %>
                                                             <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%
                                                             kol2_belobTot = kol2_belobTot + kol2_belob  
                                                             end if


                                                              case 2
                                                             
                                                             if cint(viskol1) = 1 then
                                                             %>
                                                              <td align=right style="white-space:nowrap;">
                                                              <%if kol3_belob <> 0 then %>
                                                             <%=formatnumber(kol3_belob,0) %>
                                                              <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%
                                                             kol3_belobTot = kol3_belobTot + kol3_belob  
                                                             end if
                                                             
                                                             case 3



                                                             case 4
                                                             if cint(viskol2) = 1 then

                                                             '**** Real timer / Omsætning / Kost / Faktureret / Effektiv tp 
                                                             %>
                                                             <td align=right style="white-space:nowrap;">
                                                             <%if kol9_belob <> 0 then %>
                                                             <%=formatnumber(kol9_belob,0) %>
                                                              <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%kol9_belobTot = kol9_belobTot + kol9_belob %>
                                                             
                                                             <td align=right style="white-space:nowrap;">
                                                             <%if kol4_belob <> 0 then %>
                                                             <%=formatnumber(kol4_belob,0) %>
                                                              <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%kol4_belobTot = kol4_belobTot + kol4_belob %>
                                                             
                                                             <td align=right style="white-space:nowrap;">
                                                             <%if kol5_belob <> 0 then %>
                                                             <%=formatnumber(kol5_belob,0) %>
                                                              <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%kol5_belobTot = kol5_belobTot + kol5_belob %>

                                                             

                                                             <%
                                                             
                                                             end if
                                                             
                                                             case 5 
                                                              '***** SUm og Effektiv tp, lukkede job ***'
                                                              if cint(viskol2) = 1 then
                                                              %>
                                                              <td align=right style="white-space:nowrap;">
                                                              <%if kol6_belob <> 0 then %>
                                                              <%=formatnumber(kol6_belob,0) %>
                                                               <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                              <%kol6_belobTot = kol6_belobTot + kol6_belob %>

                                                             <%
                                                               sumLukkedeJobkos_fak = 0
                                                               sumLukkedeJobkos_fak = (kol6_belob - (kol5_belob))

                                                               

                                                               '*** Fratrækker salgsomkostninger, da de ikke er fratrukket for Epi
                                                             
                                                               if kol9_belob <> 0 then
                                                               faktiskTp = ((kol11_belob-kol10_belob)/kol9_belob)
                                                               else
                                                               faktiskTp = 0
                                                               end if

                                                              

                                                               
                                                               %>
                                       
                                                                <td align=right style="white-space:nowrap;"><%=formatnumber(sumLukkedeJobkos_fak, 0) %></td>

                                                               <td align=right style="white-space:nowrap;">
                                                             <%if kol10_belob <> 0 then %>
                                                             <%=formatnumber(kol10_belob,0) %>
                                                             </td>
                                                             <%kol10_belobTot = kol10_belobTot + kol10_belob 
                                                             
                                                             end if
                                                             %>


                                                                  <td align=right style="white-space:nowrap;">
                                                                  <%if faktiskTp <> 0 then %>
                                                                  <%=formatnumber(faktiskTp,0) %>
                                                                  <%else %>
                                                                  &nbsp;
                                                                  <%end if %></td>


                                                               <%

                                                           
                                                            sumLukkedeJobkos_fakTot = sumLukkedeJobkos_fakTot + (sumLukkedeJobkos_fak) 
                                                            faktiskTpTot = faktiskTpTot + (faktiskTp)
                                                            

                                                             end if

                                                             case 6

                                                                 if cint(viskol3) = 1 then
                                                                 '** Åbne job, Real.
                                                                   %>
                                                                  <td align=right style="white-space:nowrap;">
                                                                  <%if kol12_belob <> 0 then %>
                                                                  <%=formatnumber(kol12_belob,0) %>
                                                                   <%else %>
                                                                 &nbsp;
                                                                 <%end if %></td>
                                                                  <%kol12_belobTot = kol12_belobTot + kol12_belob %>

                                                                   <td align=right style="white-space:nowrap;">
                                                                   <%if kol7_belob <> 0 then %>
                                                                   <%=formatnumber(kol7_belob,0) %>
                                                                    <%else %>
                                                                 &nbsp;
                                                                 <%end if %></td>
                                                                   <%kol7_belobTot = kol7_belobTot + kol7_belob %>
                                                                 <%
                                                                 end if


                                                             case 7
                                                             '** Åbne job, Faktureret
                                                             if cint(viskol3) = 1 then
                                                             %>
                                                              <td align=right style="white-space:nowrap;">
                                                              <%if kol8_belob <> 0 then %>
                                                              <%=formatnumber(kol8_belob,0) %>
                                                               <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                              <%kol8_belobTot = kol8_belobTot + kol8_belob %>
                                                              
                                                              <% 
                                                               sumAbneJobkos_fak = 0
                                                               sumAbneJobkos_fak = (kol8_belob - (kol7_belob)) %>
                                                              
                                                              <td align=right style="white-space:nowrap;"><%=formatnumber(sumAbneJobkos_fak,0) %></td>

                                                             <%
                                                             sumAbneJobkos_fakTot = sumAbneJobkos_fakTot + (sumAbneJobkos_fak)


                                                             '**** Salgsomkostninger Åbne job ****
                                                              %>
                                                             <td align=right style="white-space:nowrap;">
                                                             <%if kol13_belob <> 0 then %>
                                                             <%=formatnumber(kol13_belob,0) %>
                                                              <%else %>
                                                             &nbsp;
                                                             <%end if %></td>
                                                             <%kol13_belobTot = kol13_belobTot + kol13_belob %>
                                                             
                                                             <%


                                                             '** Faktisk tp åbne job ****
                                                              if kol12_belob <> 0 then
                                                               faktiskTpAbne = ((kol8_belobTot-kol13_belob)/kol12_belob)
                                                               else
                                                               faktiskTpAbne = 0
                                                               end if

                                                            

                                                             %> <td align=right style="white-space:nowrap;"><%=formatnumber(faktiskTpAbne,0) %></td><%

                                                             faktiskTpAbneTot = faktiskTpAbneTot + faktiskTpAbne 

                                                              end if

                                                             case 8

                                                             
                                                           

                                                             case 9

                                                            
                                                             
                                                             

                                                             case 10

                                                               if cint(viskol6) = 1 then

                                                               if kol14_belob <> 0 then
                                                               effektivTpKvo = kol14_belob '(kol6_belob/kol5_belob)
                                                               else
                                                               effektivTpKvo = 0
                                                               end if %>

                                                                  <td align=right style="white-space:nowrap;">
                                                                  <%if effektivTpKvo <> 0 then %>
                                                                  <%=formatnumber(effektivTpKvo,2) %> 
                                                                  <%else %>
                                                                  &nbsp;
                                                                  <%end if %></td>



                                                             <%
                                                             end if

                                                             effektivTpKvoTot = effektivTpKvoTot + (effektivTpKvo)
                                                             end select %>

                                                             
                                                            
                                                   
                                                   
                                                   
                                                             
                                                       

                                                            <%'formatnumber(medarbBelob(oRec2("mid"),m),0) %> 

                                                            <%
                                                            

                                                else



                                                         select case m
                                                             case 0,1,2
                                                             
                                                             if cint(viskol1) = 1 then
                                                             strExport = strExport & formatnumber(medarbBelob(oRec2("mid"),m),0) & ";"
                                                             end if
                                                             
                                                             case 4
                                                             if cint(viskol2) = 1 then
                                                             strExport = strExport & formatnumber(kol9_belob,0) & ";"
                                                             strExport = strExport & formatnumber(kol4_belob,0) & ";"
                                                             strExport = strExport & formatnumber(kol5_belob,0) & ";"
                                                            
                                                             end if

                                                             case 5 
                                                             
                                                              if cint(viskol2) = 1 then
                                                               strExport = strExport & formatnumber(kol6_belob,0) & ";"
                                                            
                                                               sumLukkedeJobkos_fak = 0
                                                               sumLukkedeJobkos_fak = (kol6_belob - (kol5_belob))

                                                              
                                                              '***** Faktisk tp lukkede job
                                                               if kol9_belob <> 0 then
                                                               faktiskTp = ((kol11_belob-kol10_belob)/kol9_belob)
                                                               else
                                                               faktiskTp = 0
                                                               end if

                                                              

                                                               strExport = strExport & formatnumber(sumLukkedeJobkos_fak,0) & ";"
                                                               strExport = strExport & formatnumber(kol10_belob,0) & ";"
                                                               strExport = strExport & formatnumber(faktiskTp,0) & ";"
                                                              
                                                            
                                                              
                                                             end if

                                                             case 6

                                                                 if cint(viskol3) = 1 then
                                                                 strExport = strExport & formatnumber(kol12_belob,0) & ";"
                                                                 strExport = strExport & formatnumber(kol7_belob,0) & ";"
                                                                 end if

                                                             
                                                             case 7
                                                                 if cint(viskol3) = 1 then
                                                                  strExport = strExport & formatnumber(kol8_belob,0) & ";"
                                                            
                                                                   sumAbneJobkos_fak = 0
                                                                   sumAbneJobkos_fak = (kol8_belob - (kol7_belob))
                                                                 
                                                                 strExport = strExport & formatnumber(sumAbneJobkos_fak,0) & ";"
                                                                 strExport = strExport & formatnumber(kol13_belob,0) & ";"

                                                                  '** Faktisk tp åbne job ****
                                                                if kol12_belob <> 0 then
                                                               faktiskTpAbne = ((kol8_belobTot-kol13_belob)/kol12_belob)
                                                               else
                                                               faktiskTpAbne = 0
                                                               end if

                                                                strExport = strExport & formatnumber(faktiskTpAbne,0) & ";"

                                                                 end if

                                                             case 8
                                                                  
                                                                  
                                                               
                                                                   

                                                             case 9
                                                                
                                                            case 10
                                                               
                                                               
                                                               
                                                               if cint(viskol6) = 1 then


                                                               if kol14_belob <> 0 then
                                                               effektivTpKvo = kol14_belob
                                                               else
                                                               effektivTpKvo = 0
                                                               end if
                                                               
                                                               strExport = strExport & formatnumber(effektivTpKvo,2) & ";"

                                                               end if

                                                             end select
                                                
                                          
                                                '** Export ***'
                                                'if medarbBelob(oRec2("mid"),m) <> 0 then
                                                'strExport = strExport & formatnumber(medarbBelob(oRec2("mid"),m),0) & ";"
                                                '''Response.Write "strExport: m: "& m &", " & strExport & "<br>"
                                                'else
                                                '    if media = "chart" then
                                                '    strExport = strExport & "0;"
                                                '    else
                                                '    strExport = strExport & ";"
                                                '    end if
                                                'end if

                                                ''**** Graf 2 *****'
                                                'if medarbBelob(oRec2("mid"),m) <> 0 then
                                                ' strExportA(mids,m) = formatnumber(medarbBelob(oRec2("mid"),m),0)
                                                ' else
                                               'strExportA(mids,m) = 0
                                               'end if

                                          
                                         
                                         end if
                                
                                
                                         

                                  
                                            


        
                             


                            next

   
    strExport = strExport & "xx99123sy#z" 
     

    

     if media <> "export" AND media <> "chart" then 
     
     
      '** Grundlag ***'
     if cint(viskol4) = 1 then
     %>
     <td width="200" valign=top><a href="#" id="glag_<%=oRec2("mid") %>" class="glag">+ Se grundlag</a><div id="divglag_<%=oRec2("mid") %>" style="font-size:9px; color:#999999; width:200px; visibility:hidden; display:none;">
     
     <%for m = 0 to 10 
     
      select case m
     %>
     
     
     <%case 0%>
     <br /><br /><br /><span style="color:#5582d2;"><b>Budget, tilbud</b></span><br />
     <%case 2%>
     <br /><br /><br /><span style="color:#5582d2;"><b>Budget, Job (WIP)</b></span><br />
     <%case 3%>
     <br /><br /><br /><span style="color:#5582d2;"><b>Lukkede job, salgsomkostninger</b></span><br />
     <%case 4 %>
     <br /><br /><br /><span style="color:#5582d2;"><b>Lukkede job, realiseret</b></span><br />
     <%case 6 %>
     <br /><br /><br /><span style="color:#5582d2;"><b>Åbne job, realiseret</b></span><br />
     <%case 8 %>
     
     <%case 9 %>
     <br /><br /><br /><span style="color:#5582d2;"><b>Åbne job, salgsomkostninger</b></span><br />
    
     <%case 10 %>
     <br /><br /><br /><span style="color:#5582d2;"><b>Lukkede job, Faktor</b></span><br />
     <%end select %>

     <%=medarbGrlag(oRec2("mid"),m) %>
     <%next %>

     <!--
     <br /><br />
     <b>Samlede salgs. omkost.:</b><br /> <%=matkobspris %>
     -->
     
     </div></td>
     <%end if %>

    </tr>
    <%end if


    
    mids = mids + 1
    oRec2.movenext
    wend
    oRec2.close



    if mids > 0 then
    mids = mids
    else
    mids = 1
    end if

    '********** totaler *******'


        if media <> "export" AND media <> "chart" then
        %><tr><td><b>Total:</b></td><%

          if cint(viskol1) = 1 then
          %>
          <td align=right><%=formatnumber(kol1_belobTot,0)%></td>
          <td align=right><%=formatnumber(kol2_belobTot,0)%></td>
          <td align=right><%=formatnumber(kol3_belobTot,0)%></td>
          <%
          end if   
          
          if cint(viskol2) = 1 then
          %>
          <td align=right><%=formatnumber(kol9_belobTot,0)%></td>
          <td align=right><%=formatnumber(kol4_belobTot,0)%></td>
          <td align=right><%=formatnumber(kol5_belobTot,0)%></td>
          
          <td align=right><%=formatnumber(kol6_belobTot,0)%></td>

          <td align=right><%=formatnumber(sumLukkedeJobkos_fakTot,0)%></td>
           <td align=right><%=formatnumber(kol10_belobTot,0)%></td>
          <td align=right><%=formatnumber(faktiskTpTot/mids,0) %></td>
       
          <%end if                                                  
                                                           
         
          if cint(viskol3) = 1 then
          %>
           <td align=right><%=formatnumber(kol12_belobTot,0)%></td>
           <td align=right><%=formatnumber(kol7_belobTot,0)%></td>
           <td align=right><%=formatnumber(kol8_belobTot,0)%></td>

           <td align=right><%=formatnumber(sumAbneJobkos_fakTot,0)%></td>
           <td align=right><%=formatnumber(kol13_belobTot,0)%></td>
           <td align=right><%=formatnumber(faktiskTpAbneTot/mids,0) %></td>


          <% 
         end if
        
        

         if cint(viskol6) = 1 then
          %>
             <td align=right>&nbsp;</td>
          
           <%end if

          '** Grundlag ***'
     if cint(viskol4) = 1 then

         %>
       
         <td>&nbsp;</td>
         <%end if %>
         </tr>
    
    
        </table>
        <br /><br />
        *) Work in progress. (restestimat)<br />
       
       Job med prioitet -1 (interne) bliver ikke medtaget i statistikken.<br />&nbsp;
        </div>

        <%end if 'export %>
	


    <%

    'end if

  


    '******************* Eksport **************************' 

                            if media = "export" OR media = "chart" then

    
                              call TimeOutVersion() 

                                
                                ekspTxt = strExport
	                            ekspTxt = replace(ekspTxt, ";xx99123sy#z", vbcrlf)


                                'Response.Write ekspTxt
                                'Response.end


                                for e = 0 TO mids-1 'UBOUND(strExportA)
                                ekspTxtAval0 = ekspTxtAval0 & strExportA(e,0) & ";"
                                ekspTxtAval1 = ekspTxtAval1 & strExportA(e,1) & ";"
                                ekspTxtAval2 = ekspTxtAval2 & strExportA(e,2) & ";"
                                ekspTxtAval3 = ekspTxtAval3 & strExportA(e,3) & ";"
                                ekspTxtAval4 = ekspTxtAval4 & strExportA(e,4) & ";"
                                ekspTxtAval5 = ekspTxtAval5 & strExportA(e,5) & ";"
                                next

                                 

                                ekspTxtAval0 = "Medarb./kategorier;"& ekspTxtAval0 & "xx99123sy#z"
                                ekspTxtAval1 = "Tilbud : Andel % : "&toplinie_txt &";"& ekspTxtAval1 & "xx99123sy#z"
                                ekspTxtAval2 = "Tilbud * Sandsyn. % : Andel : "&toplinie_txt&";"& ekspTxtAval2 & "xx99123sy#z"
                                ekspTxtAval3 = "Aktive job : WIP* (restestimat) : Andel : "&toplinie_txt &";"& ekspTxtAval3 & "xx99123sy#z"
                                ekspTxtAval4 = "Lukkede job : Andel ("&toplinie_txt&") : Faktureret oms.;"& ekspTxtAval4 & "xx99123sy#z"
                                ekspTxtAval5 = "Alle job : Andel ("&toplinie_txt&") : Faktureret oms.;"& ekspTxtAval5 & "xx99123sy#z"

	                            ekspTxtAval = ekspTxtAval0 & ekspTxtAval1 & ekspTxtAval2 & ekspTxtAval3 & ekspTxtAval4 & ekspTxtAval5
                                ekspTxtAval = replace(ekspTxtAval, ";xx99123sy#z", vbcrlf)
	
                                'datointerval = request("datointerval")
	                            

                               
                                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
                               
                               
	
	                           
	
				                            Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\saleandvalue.asp" then
					                            Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                            Set objNewFile = nothing
					                            Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                            else
					                            Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                            Set objNewFile = nothing
					                            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                            end if
				
				
				
				                            file = "saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"




                                            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\saleandvalue.asp" then
					                            Set objNewFileA = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"_a.csv", True, False)
					                            Set objNewFileA = nothing
					                            Set objFA = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"_a.csv", 8)
				                            else
					                            Set objNewFileA = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"_a.csv", True, False)
					                            Set objNewFileA = nothing
					                            Set objFA = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"_a.csv", 8)
				                            end if

                                            fileA = "saleandvalueexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"_a.csv"

                                             %>
                                            <form>
                                            <input type="hidden" value="<%=file %>" name="FM_filename" id="FM_filename" />
                                            <input type="hidden" value="<%=fileA %>" name="FM_filenameA" id="FM_filenameA" />
                                            </form>
                                            <%
				
				                                  
				                                

                                              
				                                
				                            
                                            
                                               
                                               '**** Eksport fil, kolonne overskrifter ***

                                               strOskrifter = ";"

                                               if cint(viskol1) = 1 then
                                               strOskrifter = strOskrifter & ";;;"
                                               end if

                                               if cint(viskol2) = 1 then
                                               strOskrifter = strOskrifter & "Lukkede job;;;;;;;"
                                               end if

                                               if cint(viskol3) = 1 then
                                               strOskrifter = strOskrifter & "Åbne Job;;;;;;"
                                               end if

                                              

                                               if cint(viskol6) = 1 then
                                               strOskrifter = strOskrifter & "Effektiv timepris;"
                                               end if
                                               
                                               
                                               strOskrifter = strOskrifter & vbcrlf


                                               strOskrifter = strOskrifter & "Medarbejder;"
                                               if cint(viskol1) = 1 then
				                               strOskrifter = strOskrifter &"Tilbud : Andel % : "&toplinie_txt &" "& basisValISO &""_
                                               &";Tilbud * Sandsyn. % : Andel : "&toplinie_txt &" "& basisValISO & ""_
                                               &";Aktive job : WIP* (restestimat) : Andel : "&toplinie_txt &" "& basisValISO &";"
                                               end if

                                               if cint(viskol2) = 1 then
                                               strOskrifter = strOskrifter &"Real. timer;Real. Omsætning "& basisValISO &";Real. Omkost. "& basisValISO &";Faktureret "& basisValISO &";Sum "& basisValISO &";Salgsomk.;Faktisk timepris "& basisValISO &";"
                                               end if
                                               
                                               if cint(viskol3) = 1 then
                                               strOskrifter = strOskrifter &"Real. timer;Real. Omkost. "& basisValISO &"; Faktureret "& basisValISO &"; Sum "& basisValISO &";Salgsomk.;Faktisk timepris "& basisValISO &";"
                                               end if

                                            

                                                   if cint(viskol6) = 1 then
                                                   strOskrifter = strOskrifter &"Faktor, lukkede job i periode, uanset andel;"
                                                   end if

                                                if media = "export" then
                                               objF.writeLine("Periode afgrænsning: "& formatdatetime(stDatoShow, 1) &" til "& formatdatetime(slDatoShow, 1) & " ("& toplinie_txt &")"& vbcrlf)
                                                end if  

                                            objF.WriteLine(strOskrifter & chr(013))
                                            objF.WriteLine(ekspTxt)
				                            objF.close

                                            objFA.WriteLine(ekspTxtAval)
				                            objFA.close
				
				                            if media = "chart" then%>

                                            
                                            <script src="inc/sale_jav.js"></script>

                                              <%if mids <= 10 then
                                             ct2hgt = "400"
                                             end if %>

                                              <%if mids > 10 AND mids <= 20 then
                                             ct2hgt = "600"
                                             end if %>

                                             <%if mids > 20 AND mids <= 40 then
                                             ct2hgt = "800"
                                             end if %>

                                             <%if mids > 40 AND mids <= 60 then
                                             ct2hgt = "1000"
                                             end if %>

                                               <%if mids > 60 AND mids <= 100 then
                                             ct2hgt = "1400"
                                             end if %>

                                               <%if mids > 100 then
                                             ct2hgt = "2000"
                                             end if %>

                                              <%if mids > 200 then
                                             ct2hgt = "3000"
                                             end if %>

                                             <div id="container" style="width: 1020px; height: <%=ct2hgt%>px; border:0px;"></div>
                                             <br /><br /><br /><br />

                                           

                                             <div id="container2" style="width: 1020px; height: <%=ct2hgt%>px; border:0px;"></div>

                                            <%else %>
				
	                                        <table border=0 cellspacing=1 cellpadding=0 width="200">
	                                        <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                                        <img src="../ill/outzource_logo_200.gif" />
	                                        </td>
	                                        </tr>
	                                        <tr>
	                                        <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	                                        <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Din CSV. fil er klar >></a>
	                                        </td></tr>
	                                        </table>
	            
	          
	            
	                                        <%end if
                
                
                                            Response.end
	                                        'Response.redirect "../inc/log/data/"& file &""	
				



                            end if






                            if media <> "print" AND media <> "export" AND media <> "chart" then

                        ptop = 0
                        pleft = 830
                        pwdt = 140

                        call eksportogprint(ptop,pleft,pwdt)
                        lnk = "&FM_kunde="&kundeid&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&FM_progrp="&progrp&"&seomsfor="&seomsfor&"&FM_start_mrd="&strMrd&""_
                        &"&interval="&interval&"&FM_toplinie_db="&toplinie_db&"&FM_virksomhedsandel="&virkAndel&"&FM_kol_1="&viskol1&"&FM_kol_2="&viskol2&"&FM_kol_3="&viskol3&"&FM_kol_4="&viskol4&"&FM_kol_5="&viskol5&"&FM_kol_6="&viskol6

                        %>

        
                              <tr>
                                <td align=center>
                                <a href="#" onclick="Javascript:window.open('saleandvalue.asp?media=export<%=lnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
                                </td><td><a href="#" onclick="Javascript:window.open('saleandvalue.asp?media=export<%=lnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport</a></td>
                               </tr>
                            <tr>
                           <td align=center><a href="saleandvalue.asp?media=print<%=lnk%>" target="_blank"  class='rmenu'>
                           &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
                            </td><td><a href="saleandvalue.asp?media=print<%=lnk%>" target="_blank" class="rmenu">Print version</a></td>
                           </tr>
                           <!--
                           <tr>
                            <td align=center><a href="saleandvalue.asp?media=chart<%=lnk%>" target="_blank"  class='rmenu'>
                           &nbsp;<img src="../ill/chart_column.png" border="0" alt="" /></a>
                            </td><td><a href="saleandvalue.asp?media=chart<%=lnk%>" target="_blank" class="rmenu">Graf</a></td>
                           

                           </tr>
                           -->
   
	
                           </table>
                        </div>
                        <%else%>

                        <% 
                        Response.Write("<script language=""JavaScript"">window.print();</script>")
                        %>
                        <%end if%>



	<br>
	<br>
	
	&nbsp;
	
    
	
	</div><!-- side div -->
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
