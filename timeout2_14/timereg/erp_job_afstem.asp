<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
		
    menu = "erp"
	'print = request("print")
	media = request("media")
    	
		
    if media <> "print" then%>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <SCRIPT language=javascript src="inc/job_afstem_jav.js"></script>
    
  <%call menu_2014() %>
    


    <%else%>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
    <%end if
		
	'*** Finder akttyper der er fakturerbare, ikke fakbare / tælles med i timereg ***	
	call akttyper2009(2)	

           
     '*** Joblog / Kunde niveau / jobniveau / medarb niveau **'
     public medarbTimerKunde, medarbTimerFBKunde, medarbOmsKunde, medarbValutaKunde, medarbFaknrs
     redim medarbTimerKunde(3000), medarbTimerFBKunde(3000), medarbOmsKunde(3000), medarbValutaKunde(3000), medarbFaknrs(3000) 
            
     function joblogtimer(visjoblogPer, stDato, slDato, kid, jobnr, medid, faktype)
            

           

             if jobnr <> 0 then
                jidSQLkri = " AND tjobnr = '" & jobnr & "'"
             else
                jidSQLkri = " AND tjobnr <> '0'"
             end if

           

             if medid <> -1 then
                midSQLkri = " AND tmnr = " & medid
                'grpBYmed = ", tmnr"
             else
                midSQLkri = " AND tmnr <> 0"
                'grpBYmed = ""
             end if

			'realtimer(x) = 0
			
			'**** Alle timer *****
            strSQL5 = "SELECT timer AS realtimer, tmnr, t2.kurs, t2.valuta, v.valutakode, j.fastpris, j.jobTpris, tjobnr, tdato, "_
            &" (budgettimer+ikkebudgettimer) AS budgettimer, timepris, mid, mnr, t2.tfaktim FROM "_
			&" timer t2 "_
            &" LEFT JOIN medarbejdere AS m ON (m.mid = t2.tmnr)"_
            &" LEFT JOIN valutaer v ON (v.id = 1) "_
            &" LEFT JOIN job j ON (j.jobnr = t2.tjobnr) "_
            &" WHERE t2.tknr = "& kid &" AND "_
			&" t2.tdato BETWEEN '"& stDato &"' AND '" & slDato &"' AND ("& replace(aty_sql_realHours, "tfak", "t2.tfak") &") "& jidSQLkri &" "& midSQLkri 

			
			
            'Response.Write strSQL5 & "<br><br>"
            'Response.flush

			oRec5.open strSQL5, oConn, 3
			while not oRec5.EOF 
			
            call akttyper2009Prop(oRec5("tfaktim"))

            if cint(aty_fakbar) = 1 then 'fakturerbar


             faktimertimepris = 0
            '** Tager højde for kreditnotaer
            if cint(faktype) <> 1 then
            medarbTimerFBKunde(oRec5("mid")) = medarbTimerFBKunde(oRec5("mid")) + oRec5("realtimer")
			fbtimer(x) = fbtimer(x) + oRec5("realtimer")
            else
            medarbTimerFBKunde(oRec5("mid")) = medarbTimerFBKunde(oRec5("mid")) - oRec5("realtimer")
			fbtimer(x) = fbtimer(x) - oRec5("realtimer")
            end if

			
                '** Oms
            

                if oRec5("fastpris") <> "1" then 
					if len(oRec5("realtimer")) <> 0 AND oRec5("realtimer") <> 0 then
					faktimertimepris = oRec5("timepris")/1
					else
					faktimertimepris = 0
					end if
				else
					if len(oRec5("budgettimer")) <> 0 ANd oRec5("budgettimer") <> 0 then
				    faktimertimepris = (oRec5("jobTpris") / oRec5("budgettimer"))
					else
					faktimertimepris = oRec5("jobTpris")/1
					end if
				end if
									
				
									
				call beregnValuta(faktimertimepris,oRec5("kurs"),100)
				faktimertimepris = valBelobBeregnet

                '** Tager højde for kreditnotaer
                'if oRec5("fbtimer") <> 0 then

                if cint(faktype) <> 1 then
                medarbOmsKunde(oRec5("mid")) = medarbOmsKunde(oRec5("mid")) + (oRec5("realtimer") * faktimertimepris)
                timerOmsKunde(x) = timerOmsKunde(x) + (oRec5("realtimer") * faktimertimepris)
                else
                medarbOmsKunde(oRec5("mid")) = medarbOmsKunde(oRec5("mid")) - (oRec5("realtimer") * faktimertimepris)
                timerOmsKunde(x) = timerOmsKunde(x) - (oRec5("realtimer") * faktimertimepris)
                end if

                'end if

				medarbValutaKunde(oRec5("mid")) = oRec5("valutakode")		

               


            end if



            '*** Alle real timer der tæller med idagligt timeregnskab ***
            '** Tager højde for kreditnotaer
            if cint(faktype) <> 1 then
            medarbTimerKunde(oRec5("mid")) = medarbTimerKunde(oRec5("mid")) + oRec5("realtimer")
			realtimer(x) = realtimer(x) + oRec5("realtimer")
            else
            medarbTimerKunde(oRec5("mid")) = medarbTimerKunde(oRec5("mid")) - oRec5("realtimer")
			realtimer(x) = realtimer(x) - oRec5("realtimer")
            end if

            response.flush
			
            oRec5.movenext
			wend 
			oRec5.close
			
			
		
            
           
           
        end function


     
      public medarbFakTimerKunde, medarbFakBelobKunde
      redim medarbFakTimerKunde(1000), medarbFakBelobKunde(1000)
	function xmedarbFakKunde(medid, meNr, stDato, slDato, kid)
				

            if kid <> -1 then
		    kundeKriID = " AND fakadr = " & kid 
			else
			kundeKriID = " AND fakadr <> " & kid 
			end if

				
				strSQL2 = "SELECT SUM(fm.fak) AS medfaktimer, SUM(fm.beloeb) AS medfakbel, "_
				&" f.faktype, f.kurs AS fakkurs, f.valuta, fm.valuta, fm.kurs AS fmskurs "_
				&" FROM fakturaer f "_
				&" LEFT JOIN fak_med_spec fm ON (fm.fakid = f.fid AND fm.mid = "& medid &") WHERE "_
				&" (fakdato BETWEEN '"& stDato &"' AND '" & slDato &"' "& kundeKriID &") AND fm.enhedsang = 0 "& interneSQLkri &" AND shadowcopy = 0 "_
				&" GROUP BY f.faktype, fm.mid, fm.valuta, fm.kurs"
				
				'Response.write strSQL2 &"<br><br>"
				'Response.flush
				
				medFakTimerKundeTot = 0
                medFakBelKundeTot = 0
				'medarbFakTimerKunde(meNr) = 0
				'medarbFakBelobKunde(meNr) = 0

				oRec2.open strSQL2, oConn, 3 
				while not oRec2.EOF 
				
              
				
				'** Beløb i FMS tabel er korrigeret for KURS ***
				
		        
				call beregnValuta(oRec2("medfakbel"),oRec2("fakkurs"),100) '*** TIL DKK
			    valBelobBeregnet = valBelobBeregnet
				
				if oRec2("faktype") <> 1 then
				medFakTimerKundeTot = medFakTimerKundeTot + oRec2("medfaktimer")
				medFakBelKundeTot = medFakBelKundeTot + valBelobBeregnet
                else
				medFakTimerKundeTot = medFakTimerKundeTot - (oRec2("medfaktimer"))
				medFakBelKundeTot = medFakBelKundeTot - (valBelobBeregnet)
                end if
				
				oRec2.movenext
				wend
				oRec2.close 
			    

			    medarbFakTimerKunde(m) = medFakTimerKundeTot
			    medarbFakBelobKunde(m) = medFakBelKundeTot

	end function

		
	'public totfaktimer, totfakbelob, totmedtimer, totmedbel 
	public totMtimer, totMFBTimer, totMoms, totMfaktimer, totMfakBelob

             totMtimer = 0
            totMFBTimer = 0
            totMoms = 0
            totMfaktimer = 0
            totMfakBelob = 0

    function medTotalerKid()
    %>
    <br><br>
	<img src="../ill/users2.png" width="24" height="24" alt="" border="0">&nbsp;<b>Medarbejder totaler:</b>  
    <%if visjoblogPer = 1 then %>
    (kun på job der er med på fakturaer ovenfor)
    <%end if %>
    <br><br>
			<table cellpadding=2 cellspacing=0 border=0 width=100%>
			<tr bgcolor="#D6DFf5">
				<td bgcolor="#5582d2" valign=bottom class=lille style="padding-left:2px; border-right:1px #FFFFFF solid;">Navn</td>
				<td bgcolor="#5582d2" valign=bottom class=lille style="border-right:1px #FFFFFF solid;" align=right>Registreret</td>
                <td bgcolor="#5582d2" valign=bottom class=lille style="border-right:1px #FFFFFF solid;" align=right>Heraf fakturerbare</td>
                <td bgcolor="#5582d2" valign=bottom class=lille style="border-right:1px #FFFFFF solid;" align=right>Omsætning</td>
				<td valign=bottom style="padding-left:2px; border-right:1px #FFFFFF solid;" align=right class=lille>Timer faktureret<br />
				(stk, enhed og km bliver ikke medregnet)</td>
                <td valign=bottom style="padding-left:2px; border-right:1px #FFFFFF solid;" class=lille align=right>Beløb</td>
                <td valign=bottom style="border-right:1px #FFFFFF solid; width:300px; padding-left:10px;" class=lille>Fakturaer</td>
			</tr>

            <%

          
            '** m = 1 pga NULL værdier i mid felt, hvor mid så bliver sat til 0
            l = 0 
            for m = 1 to 1000

            if medarbTimerKunde(m) <> 0 OR medarbTimerFBKunde(m) <> 0 OR medarbFakTimerKunde(m) <> 0 OR medarbFakBelobKunde(m) <> 0 then
            l = l + 1

            select case right(l,1)
            case 0,2,4,6,8
            bgt = "#FFFFFF"
            case else
            bgt = "#EFf3FF"
            end select

            %>
            <tr bgcolor="<%=bgt%>">
            <td class=lille style="border-bottom:1px #cccccc solid; white-space:nowrap;" valign=top>
            <%call meStamdata(m) %>
            <b><%=meTxt %></b>
            </td>
             <td align=right class=lille style="border-bottom:1px #cccccc solid; white-space:nowrap;" valign=top><%=formatnumber(medarbTimerKunde(m), 2)  %> t.</td>
             <td align=right class=lille style="border-bottom:1px #cccccc solid; white-space:nowrap;" valign=top><%=formatnumber(medarbTimerFBKunde(m), 2)  %> t.</td>
             <td align=right class=lille style="border-bottom:1px #cccccc solid; white-space:nowrap;" valign=top><%=formatnumber(medarbOmsKunde(m), 2) &" "&basisValISO %></td>

             <td align=right class=lille style="border-bottom:1px #cccccc solid; white-space:nowrap;" valign=top><%=formatnumber(medarbFakTimerKunde(m), 2) &" t." %></td>
             <td align=right class=lille style="border-bottom:1px #cccccc solid; white-space:nowrap;" valign=top><%=formatnumber(medarbFakBelobKunde(m), 2) &" "&basisValISO %></td>
             <td class=lille style="border-bottom:1px #cccccc solid; padding-left:10px;" valign=top><%=medarbFaknrs(m) %>&nbsp;</td>
             
              
            </tr>

            <%

            totMtimer = totMtimer + (medarbTimerKunde(m))
            totMFBTimer = totMFBTimer + (medarbTimerFBKunde(m))
            totMoms = totMoms + (medarbOmsKunde(m)) 
            totMfaktimer = totMfaktimer + (medarbFakTimerKunde(m))
            totMfakBelob = totMfakBelob + (medarbFakBelobKunde(m))  
            end if

            next 
            %>

            <tr bgcolor="lightpink">
            <td class=lille>
            <b>Total:</b>
            </td>
             <td align=right class=lille style="white-space:nowrap;"><%=formatnumber(totMtimer, 2)  %> t.</td>
             <td align=right class=lille style="white-space:nowrap;"><%=formatnumber(totMFBTimer, 2)  %> t.</td>
             <td align=right class=lille style="white-space:nowrap;"><%=formatnumber(totMoms, 2) &" "&basisValISO%></td>
             <td align=right class=lille style="white-space:nowrap;"><%=formatnumber(totMfaktimer, 2) %> t.</td>
             
             <td align=right class=lille style="white-space:nowrap;"><%=formatnumber(totMfakBelob, 2) &" "& basisValISO%></td>
             <td class=lille style="white-space:nowrap; padding-left:10px;">Antal fakturaer: <%=antalfak %>&nbsp;</td>
              
            </tr>

            </table>
            <br /><br /></br>&nbsp;

    <%
    end function
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	 
	
	
	
	
	'*****************************************************************************
	'** Periode afgrænsning ***
	'*****************************************************************************
	 
	'*** Periode vælger er brugt ****
	 if len(request("FM_start_dag")) <> 0 then
	'* Bruges i weekselector *
	strMrd =  request("FM_start_mrd")
	strDag  = request("FM_start_dag")
	if len(request("FM_start_aar")) > 2 then
	strAar = right(request("FM_start_aar"), 2)
	else
	strAar = request("FM_start_aar")
	end if
	
	strMrd_slut =  request("FM_slut_mrd")
	strDag_slut  =  request("FM_slut_dag")
	
	if len(request("FM_slut_aar")) > 2 then
	strAar_slut = right(request("FM_slut_aar"),2)
	else
	strAar_slut = request("FM_slut_aar")
	end if
	
	
	select case strMrd
	case 2
		if strDag > 28 then
				    select case strAar
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    strDag = 29
				    case else
				    strDag = 28
				    end select
				else
				strDag = strDag
				end if
	case 4, 6, 9, 11
		if strDag > 30 then
		strDag = 30
		else
		strDag = strDag
		end if
	end select
	
	select case strMrd_slut
	case 2
	    if strDag_slut > 28 then
				    
		    select case strAar_slut
		    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
		    strDag_slut = 29
		    case else
		    strDag_slut = 28
		    end select
				    
		else
		strDag_slut = strDag_slut
		end if
	
		
	case 4, 6, 9, 11
		if strDag_slut > 30 then
		strDag_slut = 30
		else
		strDag_slut = strDag_slut
		end if
	end select
	
	'*** Den valgte start og slut dato ***
	StrTdato = strDag &"/" & strMrd & "/20" & strAar
	StrUdato = strDag_slut &"/" & strMrd_slut & "/20" & strAar_slut
	
	else
	
		'*Brug cookie eller dagsdato?
		if len(Request.Cookies("erp_datoer")("st_dag")) <> 0 then
	
		strMrd = Request.Cookies("erp_datoer")("st_md")
		strDag = Request.Cookies("erp_datoer")("st_dag")
		strAar = Request.Cookies("erp_datoer")("st_aar") 
		strDag_slut = Request.Cookies("erp_datoer")("sl_dag")
		strMrd_slut = Request.Cookies("erp_datoer")("sl_md")
		strAar_slut = Request.Cookies("erp_datoer")("sl_aar")
		
		else
		
		
			StrTdato = date-31
			StrUdato = date 
			
			'* Bruges i weekselector *
			if month(now()) = 1 then
			strMrd = 12
			else
			strMrd = month(now()) - 1
			end if
			
			strDag = day(now())
			
			if month(now()) = 1 then
			strAar = right(year(now()) - 1, 2)
			else
			strAar = right(year(now()), 2) 
			end if
			
			strMrd_slut = month(now())
			strAar_slut = right(year(now()), 2) 
			
			if strDag > "28" then
			strDag_slut = "1"
			strMrd_slut = strMrd_slut + 1
			else
			strDag_slut = day(now())
			end if
			
			
		end if
	end if
	
	'** Indsætter cookie **
	Response.Cookies("erp_datoer")("st_dag") = strDag
	Response.Cookies("erp_datoer")("st_md") = strMrd
	Response.Cookies("erp_datoer")("st_aar") = strAar
	Response.Cookies("erp_datoer")("sl_dag") = strDag_slut
	Response.Cookies("erp_datoer")("sl_md") = strMrd_slut
	Response.Cookies("erp_datoer")("sl_aar") = strAar_slut
	Response.Cookies("erp_datoer").Expires = date + 10		
	







	
	'**** SQL datoer ***
	sqlSTdato = "20" & strAar &"/"& strMrd &"/"& strDag 
	sqlSLUTdato = "20" & strAar_slut &"/"& strMrd_slut &"/"& strDag_slut 


    if len(trim(request("jobid"))) <> 0 AND request("jobid") <> 0 then
    jobid = request("jobid")
    else
    jobid = 0
    end if


    '*** Valgt kunde ***
	if len(request("FM_kunder")) <> 0 AND request("FM_kunder") <> 0 then
	selKunde = request("FM_kunder")
	selKundeKri = " kunder.kid = " & selKunde &""
	    
	    if len(trim(request("showall"))) <> 0 AND request("showall") <> 0 then
	    showall = 1
	    else
	    showall = 0
	    end if
	    
	'showall = 0
	else
	selKunde = 0
	selKundeKri = " kunder.kid <> 0"
	showall = 1
	end if
	
	kundeid = selKunde
	
	
    '*** Joblog periode ****'
    if len(trim(request("visjoblogPer"))) <> 0 then
    visjoblogPer = request("visjoblogPer")

        if visjoblogPer = 1 then
        visjoblogPerCHK0 = ""
        visjoblogPerCHK1 = "CHECKED"
        else
        visjoblogPerCHK0 = "CHECKED"
        visjoblogPerCHK1 = ""
        end if


        response.cookies("erp")("visjoblogper") = visjoblogPer
        response.cookies("erp").expires = date + 35
    else
        
        if request.cookies("erp")("visjoblogper") <> "" then
        visjoblogPer = request.cookies("erp")("visjoblogper")

             if visjoblogPer = 1 then
            visjoblogPerCHK0 = ""
            visjoblogPerCHK1 = "CHECKED"
            else
            visjoblogPerCHK0 = "CHECKED"
            visjoblogPerCHK1 = ""
            end if


        else
        visjoblogPer = 0
        visjoblogPerCHK0 = "CHECKED"
        visjoblogPerCHK1 = ""
        end if
    end if


    if len(trim(request("FM_interne"))) <> 0 then
    visInterne = request("FM_interne") 
    else
    visInterne = 0
    end if


    visInterneCHK0 = ""
    visInterneCHK1 = ""
    visInterneCHK2 = ""

    select case visInterne
    case 0
    visInterneCHK0 = "SELECTED"
    interneSQLkri = " AND f.medregnikkeioms = 0 "
    interneTxt = "Viser ikke interne"
    case 1
    visInterneCHK1 = "SELECTED"
    interneSQLkri = ""
    interneTxt = "Viser både interne og eksterne"
    case 2
    visInterneCHK2 = "SELECTED"
    interneSQLkri = " AND (f.medregnikkeioms = 1 OR f.medregnikkeioms = 2) "
    interneTxt = "Viser kun interne"
    end select
	
	
	if media <> "print" then
	dleft = 90
	dtop = 102

    %>

      <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td class=lille>
	<img src="../ill/outzource_logo_200.gif" /><br />
    Forventet loadtid: 2-10 sek.
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>

	</div>

    <%response.flush %>

    <%
	
	else
	
	dleft = "20"
	dtop = "20"
	end if
	%>
	
		
		<div style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px;">
		<%
        
    
	
	
		
	

 call filterheader_2013(0,0,800,pTxt)

 
    'oimg = "view_1_1.gif"
	oleft = 0
	otop = 0
	owdt = 700
	oskrift = "Afstemning job og medarbejdere (timer realiseret / faktureret)"

            
 call sideoverskrift_2014(oleft, otop, owdt, oskrift)
		
 
	

%>

<table cellspacing="0" cellpadding="0" border="0" width="100%">


<tr>
<%if media <> "print" then %>
<form action="erp_job_afstem.asp?menu=erp&showall=1&jobid=0" method="post">


    <td colspan=4><b>Kontakter (kunder):</b><br />

    <%end if %>
    
    <%
    
    strValgtKunde = "(Alle)"

		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
		
        'Response.write strSQL & "<br>"
		'Response.flush		
	
	
	if media <> "print" AND media <> "export" then
		%>
		<select name="FM_kunder" id="FM_kunder" style="width:500px;" onChange="submit();">
		<option value="0">Alle - eller vælg fra liste...</option>
		<%
	end if
				
				oRec.open strSQL, oConn, 3
				k = 0
				while not oRec.EOF
				
				
				if media <> "print" AND media <> "export" then
				
				if cint(kundeid) = cint(oRec("kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<%

                else

                if cint(kundeid) = cint(oRec("kid")) then
                strValgtKunde = oRec("Kkundenavn") &"("& oRec("Kkundenr") &")"
                end if
				
				end if
				
				k = k + 1
				oRec.movenext
				wend
				oRec.close
				
				
				if media <> "print" AND media <> "export" then
				%>
				</select>
                <%end if
    %>
    <br />&nbsp;
    </td>
</tr>

<%if media = "print" then %>
<tr>
	<td colspan=4>Kunde: <b><%=strValgtKunde %></b></td>
</tr>

<%end if %>

<%if jobid <> 0 then %>
<tr>
	<td colspan=4><b>Job:</b>
    
        <%
        strValgtJob = ""
        
        strSQLvlgJob = "SELECT jobnavn, jobnr FROM job WHERE id = "& jobid 
        oRec.open strSQLvlgJob, oConn, 3
        if not oRec.EOF then

        strValgtJob = oRec("jobnavn") & " ("& oRec("jobnr") &")"

        end if
        oRec.close
        %>

        <%=strValgtJob %><br />&nbsp;
    </td>
</tr>
<%end if %>


<%if media <> "print" then %>
<tr>
	<td colspan=4><b>Vælg periode:</b><br />&nbsp;</td>
</tr>

    <!--#include file="inc/weekselector_b.asp"-->
	<td colspan=3 style="width:200px;">&nbsp;</td>
</tr>
        <%
        if len(request("FM_nuljob")) <> 0 then
        nuljob = request("FM_nuljob")
        njChk = "CHECKED"
        else
        nuljob = 0
        njChk = ""
        end if
        %>
<tr>
	<td colspan=3><br /><b>Joblog indstillinger:</b><br />
    <input type="radio" name="visjoblogPer" id="Checkbox1" value="0" <%=visjoblogPerCHK0%>>Vis joblog for valgte periode <br /><input type="radio" name="visjoblogPer" id="Radio1" value="1" <%=visjoblogPerCHK1%>>for perioder på viste fakturaer (joblog på medarbejdere vises kun for viste fakturaer) 
    <br /><br /><input type="checkbox" name="FM_nuljob" id="FM_nuljob" value="1" <%=njChk%>>Vis "Nul" job. (Job uden registreringer)<br /><br />
    
    <b>Interne fakturaer:</b> 
        <select name="FM_interne" id="Select1" style="width:206px;" onChange="submit();">
            <option value=0 <%=visInterneCHK0%>>Vis ikke interne</option>
            <option value=1 <%=visInterneCHK1%>>Vis både interne og eksterne</option>
             <option value=2 <%=visInterneCHK2%>>Vis kun interne</option>
        </select>
    </td>
    <td align=right style="padding-right:45;" valign=bottom>
        <input id="Submit1" type="submit" value=" Kør >> " /></td>
	</tr>	
</form>

<%else %>

<tr>
	<td colspan=4>Periode: <b><%=formatdatetime(StrTdato, 1)%></b> - <b><%=formatdatetime(StrUdato, 1)%></b></td>
</tr>
<tr>
	<td colspan=4>Joblog for hele perioden / for faktura perioder: <b>
    
    <%if visjoblogPer <> 0 then %>
    Følger faktura perioder
    <%else %>
    Hele perioden
    <%end if %>
    </b></td>
</tr>
<tr>
	<td colspan=4>Interne fakturaer: <b>
    
    <%=interneTxt %>
    </b></td>
</tr>




<%end if%>

</table>
<!-- filterDiv -->
	</td></tr>
	</table>
	</div>
	
<br><br>
<%

if media <> "print" then
 if showall = 0 then%>
	        <a href="Javascript:history.back()" class=vmenu><img src="../ill/arrow_left_blue.png" alt="Tilbage" border="0"> </a> 
	       <%end if%>

<%end if%>

<%		
'*** Beskyt server mod belsatning mod langt datointerval ***
if datediff("d", sqlSTdato, sqlSLUTdato) > 730 OR datediff("d", sqlSTdato, sqlSLUTdato) < 0 then 'ca 3 md.


itop = 120
ileft = 200
iwdt = 350
idsp = ""
ivzb = "visible"
iId = "tregaktmsg1"
call sidemsgId2(itop,ileft,iwdt,iId,idsp,ivzb)
%>
<br />
			<b>Der er 2 mulige årsager til at du modtager denne fejl:</b><br><br />
			1) Det valgte datointerval er for stort. <br>
			- Vælg et datointerval på mindre end 24 måneder. (730 dage)
			<br><br />
			2) Det valgte dato-interval er negativt.
		
	        <br /><br /><br />&nbsp;
	        
	</td></tr></table>
	</div>


    <%response.end %>


<%else%>
		


		
		
		
					
		
			<%
			'**** Overblik Gand Grand total ****
            '** Denne bruges ALTID
			if showall <> 1000 then
			
            if jobid <> 0 then
            strSQLKjobLeft = " LEFT JOIN job AS j ON (j.jobknr = kunder.kid)"
            strSQLKjobWH = " AND j.id = "& jobid
            else
            strSQLKjobLeft = " LEFT JOIN job AS j ON (j.jobknr = kunder.kid)"
            strSQLKjobWH = " AND j.id <> 0 "
            end if
			
			strSQL = "SELECT kkundenr, kkundenavn, "_
			&" kunder.kid, j.jobnr FROM kunder "_
            &"" & strSQLKjobLeft &""_
			&" WHERE ("& selKundeKri &" AND ketype <> 'e') "& strSQLKjobWH &" GROUP BY kunder.kid ORDER BY kunder.kkundenavn" 
			
			'Response.write strSQl
			'Response.flush
			
			'dim ikfbtimer
			public realtimer
			public fbtimer
			dim tknavn
			dim fakbeloebKunde
			dim faktimerKunde
			dim medfaktimer
			dim medfakbel
			dim tkid
			'dim ialtregtimer
			dim tknr
			dim kundeFaknrs
            dim timerOmsKunde
            dim fakmatBeloeb, fakaktBeloeb, fakaktKorsBeloeb

			'Redim ikfbtimer(0)
			redim realtimer(350)
			Redim fbtimer(350)
			Redim tknavn(350)
			Redim faktimerKunde(350)
			Redim fakbeloebKunde(350)
			Redim medfaktimer(350)
			Redim medfakbel(350)
			Redim tkid(350)
			'Redim realtimer(350)
			Redim tknr(350)
			Redim timerOmsKunde(350)
            redim kundeFaknrs(350)
            redim fakmatBeloeb(350), fakaktBeloeb(350), fakaktKorsBeloeb(350)
			
			fakTimerKTot = 0
			fakBelKTot = 0
			medFakTimTot = 0
			medFakBelTot = 0
            timerOmsKundeTot = 0
			antalfak = 0
            kundeAntalFak = 0

            eksportFid = 0

			oRec.open strSQL, oConn, 3 
			
			x = 0
			
			while not oRec.EOF 
			
			if lastKid <> oRec("kid") then
			lastKid = oRec("kid") 
			x = x + 1
			
			realtimer(x) = 0
            fbtimer(x) = 0
			
			tkid(x) = oRec("kid")
			tknavn(x) = oRec("kkundenavn")
			tknr(x) = oRec("kkundenr")
			


            if visjoblogPer = 0 then '** Joblog Følger datointerval
                
                '*** Vis for specifikt job **'
                if jobid <> 0 then
                jl_jobnr = oRec("jobnr") 
                else
                jl_jobnr = 0
                end if

                call joblogtimer(visjoblogPer, sqlSTdato, sqlSLUTdato, tkid(x), jl_jobnr,-1,-1)
            end if
			
            if jobid <> 0 then
            jobidSQL = " AND j.id = "& jobid
            else
            jobidSQL = " AND j.id <> 0"
            end if
           
			
			'*** Faktura medab. linier ***'
			'*** Kun dem med enhed = 0 (timer) ***'
			strSQL2 = "SELECT f.faktype, "_
			&" fm.fak AS medfaktimer, fm.beloeb AS medfakbel, fm.enhedsang, "_
			&" f.kurs AS fakkurs, fm.kurs AS fmskurs, fm.mid, f.faknr, f.jobid, f.aftaleid, f.fid, j.jobnavn, j.jobnr, f.beloeb AS fakbelob, f.faktype, f.fakadr "_
            &" FROM fakturaer f "_
            &" LEFT JOIN fak_med_spec fm ON (fm.fakid = f.fid) "_
            &" LEFT JOIN job AS j ON (j.id = f.jobid)"_
			&" WHERE "_
			&" ((f.fakadr = "& oRec("kid") &" AND fakdato BETWEEN '"& sqlSTdato &"' "_
			&" AND '" & sqlSLUTdato &"') "& interneSQLkri &" AND shadowcopy <> 1) "& jobidSQL &""_
			&" ORDER BY fm.fakid, fm.kurs, fm.valuta, fm.mid"
			'&" GROUP BY f.faktype, fm.fakid, fm.kurs, fm.valuta, fm.mid"
			
			medfaktimer(x) = 0
			medfakbel(x) = 0
			nyMedarbFakbel = 0
			
			'Response.write "<br><br>"& strSQL2
			'Response.flush
			
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF 
			
            if isNull(oRec2("mid")) = false then
            useMmid = oRec2("mid")
            else
            useMmid = 0
            end if


			call beregnValuta(oRec2("medfakbel"),oRec2("fakkurs"),100)
			nyMedarbFakbel = valBelobBeregnet
			
			'Response.Write oRec2("medfakbel") & " - " & oRec2("fakkurs") & "<br>"
			
            if oRec2("enhedsang") = 0 then '** Timer

			if oRec2("faktype") <> 1 then
			medfaktimer_this = medfaktimer_this + oRec2("medfaktimer")
			medfakbel_this = medfakbel_this + nyMedarbFakbel

            '** Fakutererede timer og beløb pr. medarb.
            'call meStamdata(oRec2("mid"))
             medarbFakTimerKunde(useMmid) = medarbFakTimerKunde(useMmid) + (oRec2("medfaktimer"))
             medarbFakBelobKunde(useMmid) = medarbFakBelobKunde(useMmid) + (nyMedarbFakbel)

			else
			medfaktimer_this = medfaktimer_this -(oRec2("medfaktimer"))
			medfakbel_this = medfakbel_this -(nyMedarbFakbel)

            '** Fakutererede timer og beløb pr. medarb.
             'call meStamdata()
             medarbFakTimerKunde(useMmid) = medarbFakTimerKunde(useMmid) - (oRec2("medfaktimer"))
             medarbFakBelobKunde(useMmid) = medarbFakBelobKunde(useMmid) - (nyMedarbFakbel)

			end if

            end if
			    
               if instr(medarbFaknrs(useMmid), "<span style='color:#cccccc;'>f:</span>"& oRec2("faknr")) = 0 then

               if media <> "print" then

               
               'medarbFaknrs(useMmid) = medarbFaknrs(useMmid) & " <a href=""erp_fak_godkendt_2007.asp?id="&oRec2("fid")&"&aftid="&oRec2("aftaleid")&"&jobid="&oRec2("jobid")&""" class=""lgron"" target=""_blank""><span style='color:#cccccc;'>f:</span>"& oRec2("faknr") & "</a>"
               medarbFaknrs(useMmid) = medarbFaknrs(useMmid) & " <a href=""erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id="&oRec2("fid")&"&FM_jobonoff=j&FM_kunde="& oRec2("fakadr") &"&FM_job="&oRec2("jobid")&"&FM_aftale="& oRec2("aftaleid")&"&fromfakhist=0"" class=""lgron"" target=""_blank""><span style='color:#cccccc;'>f:</span>"& oRec2("faknr") &"</a>"
     
                else
               medarbFaknrs(useMmid) = medarbFaknrs(useMmid) & " <span style='color:#cccccc;'>f:</span>"& oRec2("faknr") 
               end if

               end if

               if instr(antalfakTxt, "f:"& oRec2("faknr")) = 0 then
               antalfak = antalfak + 1
               antalfakTxt = antalfakTxt & "f:" & oRec2("faknr")
               
               eksportFid = eksportFid &","& oRec2("fid")
               
               end if

                if instr(kundeFaknrs(x), "<span style='color:#cccccc;'>f:</span>"& oRec2("faknr")) = 0 then

                if media <> "print" then
                'kundeFaknrs(x) = kundeFaknrs(x) & " <a href=""erp_fak_godkendt_2007.asp?id="&oRec2("fid")&"&aftid="&oRec2("aftaleid")&"&jobid="&oRec2("jobid")&""" class=""lgron"" target=""_blank""><span style='color:#cccccc;'>f:</span>"& oRec2("faknr") & "</a>"
                kundeFaknrs(x) = kundeFaknrs(x) & " <a href=""erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id="&oRec2("fid")&"&FM_jobonoff=j&FM_kunde="& oRec2("fakadr") &"&FM_job="&oRec2("jobid")&"&FM_aftale="& oRec2("aftaleid")&"&fromfakhist=0"" class=""lgron"" target=""_blank""><span style='color:#cccccc;'>f:</span>"& oRec2("faknr") &"</a>"
     
                else
                kundeFaknrs(x) = kundeFaknrs(x) & " <span style='color:#cccccc;'>f:</span>"& oRec2("faknr")
                end if


                if showall <> 1 OR cint(kundeid) <> 0 then
                call beregnValuta(oRec2("fakbelob"),oRec2("fakkurs"),100)
			    nyFakbel = valBelobBeregnet

                if oRec2("faktype") <> 1 then
                nyFakbel = nyFakbel
                else
                nyFakbel = -(nyFakbel)
                end if

                if media <> "print" then
                kundeFaknrs(x) = kundeFaknrs(x) & ": "& formatnumber(nyFakbel, 2) &" "& basisValISO &"<br><a href='erp_job_afstem.asp?menu=erp&fm_kunder="&tkid(x)&"&FM_interne="&visInterne&"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&"&visjoblogPer="&visjoblogPer&"&jobid="&oRec2("jobid")&"' target=""_self"" class=""rmenu"">"& left(oRec2("jobnavn"), 25) &" ("& oRec2("jobnr") &")</a><br>"
                else
                kundeFaknrs(x) = kundeFaknrs(x) & ": "& formatnumber(nyFakbel, 2) &" "& basisValISO &"<br>"& left(oRec2("jobnavn"), 25) &" ("& oRec2("jobnr") &")<br>"
                end if

                end if
                
                kundeAntalFak = kundeAntalFak + 1
               end if

			oRec2.movenext
			wend
			oRec2.close 
			
			
			
			'*** Fak tot **'
			strSQL2 = "SELECT f.fid, f.beloeb AS fakbeloeb, f.faktype, "_
			&" f.fakadr, sum(fd.antal) AS faktimer, "_
			&" f.kurs AS fakkurs, istdato, istdato2, f.jobid, j.jobnr FROM "_
			&" fakturaer f "_
			&" LEFT JOIN faktura_det fd ON (fd.fakid = f.fid) "_
            &" LEFT JOIN job j ON (j.id = f.jobid) WHERE "_
			&" (f.fakadr = "& oRec("kid") &" AND "_
			&" fakdato BETWEEN '"& sqlSTdato &"' AND '" & sqlSLUTdato &"') "& interneSQLkri &" AND shadowcopy <> 1 "& strSQLKjobWH &" "_
			&" GROUP BY f.faktype, f.fid "
			
			'Response.Write strSQL2
			'Response.flush
			
			faktimerKunde(x) = 0
			fakbeloebKunde(x) = 0
			nyFakbel = 0
			
		    'Response.write "<br><br>"& strSQL2
			'Response.flush
			
			oRec2.open strSQL2, oConn, 3 
			while not oRec2.EOF 
			
			call beregnValuta(oRec2("fakbeloeb"),oRec2("fakkurs"),100)
			nyFakbel = valBelobBeregnet
			
			if oRec2("faktype") <> 1 then
			faktimerKunde_this = faktimerKunde_this + oRec2("faktimer")
			fakbeloebKunde_this = fakbeloebKunde_this + nyFakbel 
			else
			faktimerKunde_this = faktimerKunde_this -(oRec2("faktimer"))
			fakbeloebKunde_this = fakbeloebKunde_this -(nyFakbel) 
			end if

                        
                         if visjoblogPer = 1 AND oRec2("jobid") <> 0 then '** Joblog Følger datointerval
                         istdato = year(oRec2("istdato")) &"/"& month(oRec2("istdato")) &"/"& day(oRec2("istdato")) 
                         istdato2 = year(oRec2("istdato2")) &"/"& month(oRec2("istdato2")) &"/"& day(oRec2("istdato2"))
                            call joblogtimer(visjoblogPer, istdato, istdato2, tkid(x), oRec2("jobnr"),-1, oRec2("faktype"))
                        end if


                                        
                                        if cint(kundeid) <> 0 then

                        	            '**** Materialer udsp. ***
										strSQL5 = "SELECT sum(fms.matbeloeb * ("& replace(oRec2("fakkurs"), ",", ".") &"/100)) AS fmsbel, "_
										&" fms.matfakid FROM fak_mat_spec fms "_
										&"WHERE (fms.matfakid = "& oRec2("fid") &") GROUP BY fms.matfakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakmatBeloeb(x) = fakmatBeloeb(x) + oRec3("fmsbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakmatBeloeb(x) = fakmatBeloeb(x) - oRec3("fmsbel")
										end select
										
										
										end if
										oRec3.close
										
										
										'**** Akt sum udsp. excl. kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("fakkurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang <> 3) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktBeloeb(x) = fakaktBeloeb(x) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktBeloeb(x) = fakaktBeloeb(x) - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close
										
										
										'**** Akt sum udsp. KUN kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("fakkurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = 3) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktKorsBeloeb(x) = fakaktKorsBeloeb(x) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktKorsBeloeb(x) = fakaktKorsBeloeb(x) - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close


                                        end if

         
			
			oRec2.movenext
			wend
			oRec2.close 
			
			timerOmsKundeTot = timerOmsKundeTot + timerOmsKunde(x)
			
			faktimerKunde(x) = faktimerKunde_this
			fakbeloebKunde(x) = fakbeloebKunde_this
			medfaktimer(x) = medfaktimer_this
			medfakbel(x) = medfakbel_this
			
			faktimerKunde_this = 0
			fakbeloebKunde_this = 0
			medfaktimer_this = 0
			medfakbel_this = 0
			   
			    
			   
			end if 'LastKid <>
			
			
			'realtimer(x) = 0
			
			
			
			'** Grand Totaler ***
			regtimerTot = regtimerTot + realtimer(x) 'realtimer(x)
			fakbarregtimerTot = fakbarregtimerTot + fbtimer(x) 
		    'realtimer(x) = realtimer(x) + fbtimer(x)
			'Response.write tknavn(x) &": "& realtimer(x) &" # "& fakbarregtimerTot & "<br>"
			
			'** Grand Totaler ***
			fakTimerKTot = fakTimerKTot + faktimerKunde(x)
			fakBelKTot = fakBelKTot + fakbeloebKunde(x)
			medFakTimTot = medFakTimTot + medfaktimer(x)
			medFakBelTot = medFakBelTot + medfakbel(x)
			
			'Response.write tknavn(x) &":"& medfaktimer(x) &" -- ialt:"& medFakTimTot &"<br>"
			
			oRec.movenext
			wend
			oRec.close 
			
			
          
	
	tTop = 0
	tLeft = 0
	tWdth = 1204
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>

		
           
           
            
			
			
			<table cellspacing=0 cellpadding=2 border=0 width=100%>
			<tr bgcolor="#d6dff5">
				<td bgcolor="#5582d2" height=35 valign=bottom style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
				<td class=alt bgcolor="#5582d2" colspan=3 valign=bottom style="border-bottom:1px #CCCCCC solid;"><b>Realiserede timer</b></td>
				
				<td colspan=2 valign=bottom style="border-bottom:1px #CCCCCC solid;"><br />&nbsp;<b>Faktureret</b><br /><img src="../ill/ikon_kunder_24.png" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Kontakter</b></td>
				
                 <%'** Udspec *** 
                if cint(kundeid) <> 0 then %>
                <td colspan=3 valign=bottom style="border-bottom:1px #CCCCCC solid;">&nbsp;<b>Udspecificering</b></td>
				

                <%end if %>
                
                <td colspan=2 valign=bottom style="border-bottom:1px #CCCCCC solid;">&nbsp;<img src="../ill/users2.png" width="24" height="24" alt="" border="0">&nbsp;&nbsp;<b>Medarbejdere</b></td>
				<td valign=bottom style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
               <td valign=bottom style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
               <td valign=bottom style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
			</tr>
			<tr>
				<td bgcolor="#ffffff" valign=bottom style="border-bottom:1px #CCCCCC solid; width:220px; padding-left:2px;" class=lille>Kontakt (kontakt id)</td>
				<td bgcolor="#EFf3ff" valign=bottom style="border-bottom:1px #CCCCCC solid; width:60px;" class=lille align=right>Reg. timer ialt</td>
				<td bgcolor="#d6dff5" valign=bottom style="border-bottom:1px #CCCCCC solid; width:60px;" class=lille align=right>Heraf reg. <br /> fakturerbare timer</td>
                <td bgcolor="#EFf3ff" valign=bottom style="border-bottom:1px #CCCCCC solid; width:100px;" class=lille align=right>Omsætning</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:80px;" class=lille>Tot. fak. timer incl. "herreløse" (sum-akt.)
				(stk, enhed og km bliver ikke medregnet)</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:100px;" class=lille>Tot. faktura beløb</td>

                <%'** Udspec *** 
                if cint(kundeid) <> 0 then %>
                <td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:100px;" class=lille>Heraf materiale beløb</td>
                <td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:100px;" class=lille>Heraf sum-akt. beløb (excl. kørsel)</td>
                <td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:100px;" class=lille>Heraf Kørsels akt.</td>

                <%end if %>

				<td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:60px;" class=lille>Medarb. fak. timer
				<br />(stk, enhed og km bliver ikke medregnet)</td>
				<td bgcolor="#ffffff" valign=bottom align=right style="border-bottom:1px #CCCCCC solid; width:100px;" class=lille>Medarb. faktura beløb</td>
				 <td bgcolor="#ffffff" valign=bottom style="border-bottom:1px #CCCCCC solid; width:250px; padding-left:10px;" class=lille>Fakturaer</td>
                 <td style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
                 <td style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
			</tr>
			<%
			for x = 1 to x - 0
			
				if nuljob = 1 OR (realtimer(x) <> 0 OR fbtimer(x) <> 0 OR faktimerKunde(x) <> 0 OR fakbeloebKunde(x) <> 0 OR medfaktimer(x) <> 0 OR medfakbel(x) <> 0) then %>
				<tr>
				<td bgcolor="#ffffff" valign=top style="border-bottom:1px #CCCCCC solid; padding:2px;" class=lille>
				<%if media <> "print" then%>
				<a href="erp_job_afstem.asp?menu=erp&fm_kunder=<%=tkid(x)%>&FM_interne=<%=visInterne%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&visjoblogPer=<%=visjoblogPer%>&jobid=0" target="_self" class=rmenu>
				<%=left(tknavn(x), 25)%>&nbsp;(<%=tknr(x)%>)</a>
				<%else%>
				<%=left(tknavn(x), 25)%>&nbsp;(<%=tknr(x)%>)
				<%end if%></td>
				<td align=right valign=top bgcolor="#EFf3ff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(realtimer(x), 2)%> t.</td>
				<td align=right valign=top bgcolor="#d6dff5" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fbtimer(x), 2)%> t.</td>
                <td align=right valign=top bgcolor="#EFf3ff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(timerOmsKunde(x), 2) &" "& basisValISO %> </td>
				<td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(faktimerKunde(x), 2)%> t.</td>
				<td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakbeloebKunde(x), 2) &" "& basisValISO%></td>

                  <%'** Udspec *** 
                if cint(kundeid) <> 0 then %>
                <td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakmatBeloeb(x), 2) &" "& basisValISO %></td>
                <td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakaktBeloeb(x), 2) &" "& basisValISO %></td>
                <td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakaktKorsBeloeb(x), 2) &" "& basisValISO %></td>
                <%end if %>

				<td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(medfaktimer(x), 2)%> t.</td>
				<td align=right valign=top bgcolor="#ffffff" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(medfakbel(x), 2) &" "& basisValISO %></td>
                <td bgcolor="#ffffff" valign=top style="border-bottom:1px #CCCCCC solid; padding-left:10px;" class=lille><%=kundeFaknrs(x)%>&nbsp;</td>
                <td style="border-bottom:1px #CCCCCC solid;">&nbsp;</td>
                
				<td valign=top style="border-bottom:1px #CCCCCC solid;">
				<%
				
				if formatcurrency(fakbeloebKunde(x), 2) = formatcurrency(medfakbel(x), 2) then
				Response.write "<i><font color='limegreen'>&nbsp;V</font></i>"
				else
				Response.write "<i><font color='red'>&nbsp;-</font></i>"
				end if
				%>
                
				</td>
				</tr>
				<%
				end if
			next
			%>
			
			<tr>
			<td bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><b>Total:</b></td>
			<td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(regtimerTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakbarregtimerTot)%> t.</td>
            <td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(timerOmsKundeTot, 2) &" "& basisValISO %></td>
			<td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakTimerKTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(fakBelKTot, 2) &" "& basisValISO %> </td>
                
                <%'** Udspec *** 
                if cint(kundeid) <> 0 then %>
                <td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille>&nbsp;</td>
                <td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille>&nbsp;</td>
                <td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille>&nbsp;</td>
                <%end if %>

			<td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(medFakTimTot)%> t.</td>
			<td align=right bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille><%=formatnumber(medFakBelTot, 2) &" "& basisValISO %></td>
			<td bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid; padding-left:10px;" class=lille>Antal fakturaer: <%=kundeAntalFak %></td>
            <td bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille>&nbsp;</td>
            <td bgcolor="lightpink" style="border-bottom:1px #CCCCCC solid;" class=lille>&nbsp;</td>
			</tr>
			
			</table>
			<br><br>
			<%
			call medtotalerKid()
			
			
	else
			
		
		
		
		
		end if '*** Vis alle/Vis detaljer
		
	
		
		end if '** Over 95 dage valgt!%>
		</div><!-- tablediv-->

        <%
        	if media <> "print" then

        ptop = 0
        pleft = 840
        pwdt = 200

        call eksportogprint(ptop,pleft,pwdt)
        %>
       
           
	       
          
             <form action="erp_fakturaer_eksport_2007.asp?visning=0" method="post" target="_blank">
     <input id="Hidden2" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit4" type="submit" value="A) Detail .csv fil eksport" style="font-size:9px; width:160px;" /><br />
    <span style="color:#999999; font-size:9px;">Eksporterer alle aktivitets-, medarbejder- og materiale -linjer på de viste fakturaer, samt joblog for den valgte faktura periode. (Pivot)</span>
    </td>
    </tr>
    </form>





     <tr>
            
            <td align=center>
            <a href="erp_job_afstem.asp?media=print&fm_kunder=<%=selKunde%>&FM_interne=<%=visInterne%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&showall=<%=showall %>&visjoblogPer=<%=visjoblogPer %>&jobid=<%=jobid %>" target="_blank">
           &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
            </td><td><a href="erp_job_afstem.asp?media=print&fm_kunder=<%=selKunde%>&FM_interne=<%=visInterne%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&showall=<%=showall %>&visjoblogPer=<%=visjoblogPer %>&jobid=<%=jobid %>" class='rmenu' target="_blank">Print version</a></td>
           </tr>
      </table>


        </div>
        <%else%>

        <% 
        Response.Write("<script language=""JavaScript"">window.print();</script>")
        %>
        <%end if
        %>
        </div>	
        <br /><br />&nbsp;
<%
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
