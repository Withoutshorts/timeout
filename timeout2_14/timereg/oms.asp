<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/inc_oms.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	function luft(txt)
	%>
	<tr>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
   	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
</tr>
    <tr bgcolor="#ffffff">
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"><b><%=txt %></b></td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Total</td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
   	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
</tr>
	<%
	end function
	
	
	print = request("print")
	
	
		'*** År ***'
		if len(Request("seomsfor")) <> 0 then
		seomsfor = Request("seomsfor")
		strYear = seomsfor
		strReqAar = strYear 
		else
			'if len(request("year")) <> 0 then
			'seomsfor = year("1/1/"&request("year"))
			'strYear = request("year")
			'strReqAar = strYear
			'else
			seomsfor = year(now)
			strYear = seomsfor
			strReqAar = strYear
			'end if 
		end if
	
	
	
	
	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
		
	
		
	'*** Job og Kundeans ***
	call kundeogjobans()
	
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
    progrp = 0
	end if
	
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

    '*** Top 20 ***'
        topSel5 = ""
        topSel10 = ""
        topSel20 = ""

        if len(trim(request("FM_toplist"))) <> 0 then
        
        topList = request("FM_toplist")
        select case topList
        case 5
        topSel5 = "SELECTED"
        case 10
        topSel10 = "SELECTED"
        case 20
        topSel20 = "SELECTED"
        end select
         
	    else
        topList = 20
        topSel20 = "SELECTED"
        end if


    if len(trim(request("FM_visprkunde"))) <> 0 then
    visprkunde = request("FM_visprkunde")
    else
    visprkunde = 0
    end if

    if cint(visprkunde) = 1 then
    visprkunde1CHK = "CHECKED"
    visprkunde0CHK = ""
    else
    visprkunde1CHK = ""
    visprkunde0CHK = "CHECKED"
    end if 
	

        

        if len(request("FM_toplist_kunder_lev")) <> 0 then
        toplist_kunder_lev = request("FM_toplist_kunder_lev")
        else
        toplist_kunder_lev = 0
        end if
    

        if cint(toplist_kunder_lev) = 0 then
        toplist_kunder_lev0SEL = "SELECTED" 
        toplist_kunder_lev1SEL = ""
        else
        toplist_kunder_lev1SEL = "SELECTED"
        toplist_kunder_lev0SEL = ""
        end if


	if len(request("FM_kunde")) <> 0 AND cint(visprkunde) = 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if


	
	'*** Kundeans ***
	strKnrSQLkri = ""
	
	'*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			     kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = "jobans1 = "& intMids(m)  
			     jobAns2SQLkri = "jobans2 = "& intMids(m)
                 jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     fakmedspecSQLkri = "fms.mid = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
                 jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     fakmedspecSQLkri = fakmedspecSQLkri & " OR fms.mid = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
		    jobAnsSQLkri =  " ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "xx (" & jobAns2SQLkri &")"
			jobAns3SQLkri =  "xx (" & jobAns3SQLkri &")"
			fakmedspecSQLkri = " AND ("& fakmedspecSQLkri &")"
			
	
	'** Er key acc og kundeansvarlig valgt **?
	if cint(kundeans) = 1 then
	kundeAnsSQLKri = kundeAnsSQLKri
	else
	kundeAnsSQLKri = " Kundeans1 <> -1 AND kundeans2 <> -1 "
	end if
	
	
	'***** Valgt job eller søgt på Job ****
	'** hvis Sog = Ja
	call jobsog()

	
	'*** Aftale ****
	if len(request("FM_aftaler")) <> 0 then ' AND jobid <= 0 AND len(jobSogVal) = 0 then
	aftaleid = request("FM_aftaler")
		
		'if cint(aftaleid) > 0 then 
		'jobid = 0
		'end if
		
	else
		'if cint(jobid) > 0 OR len(jobSogVal) <> 0  then
		'aftaleid = -1
		'else
		aftaleid = 0
		'end if
	end if
	
	
	

	'**** Alle SQL kri starter med NUL records ****
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	
	media = request("media")
	
	
	call basisValutaFN()
    basisValKursUse = replace(basisValKurs, ".", ",")
	
	
	'************ slut faste filter var **************		
		
	
	
	
	
	'Response.write jobnrSQLkri
	
	
	
	
	%>

	<%if request("print") <> "j" AND media <> "export" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <script src="inc/oms_jav.js"></script>

    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:400px; width:300px; background-color:#ffffff; border:10px #cccccc solid; padding:20px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%

	'exp_loadtid = 30
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	<b>ca. 10-30 sek.</b>
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>

    <%response.flush 



        call menu_2014()
        
        else%>
	<html>
	<head>
		<title>TimeOut</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	
	<%end if%>
	
	
	<%
	thisfile = "oms"




    %>
	<!--include file="inc/stat_submenu.asp"-->
	<%
	if request("print") <> "j" AND media <> "export" then
	pleft = 90
	ptop = 102 '132
	'ptopgrafik = 348
	else
	pleft = 10
	ptop = 20
	'ptopgrafik = 90
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
	
	
	<%
	
	oimg = "ikon_stat_oms_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Omsætning"
	
    if request("print") = "j" AND media <> "export" then
        'media = "print"
	    call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	end if

	if request("print") <> "j" AND media <> "export" then 
	
	call filterheader_2013(0,0,800,oskrift)%>
    	    
	  
	    <table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="oms.asp?rdir=<%=rdir %>" method="post">
	    
	<%end if %>

	    <%call medkunderjob %>
	    
	
            
    <%if request("print") <> "j" AND media <> "export" then %>
	    
	    </td>
	    </tr>
	   
	
	
   

	<tr>
		<td valign=top><br /><b>År:</b>&nbsp;
	<select name="seomsfor" id="seomsfor" style="width:85px;" onchange="submit();">
			
			<%
			useaar = 2005
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
           <input type="radio" name="FM_visprkunde" value="0" <%=visprkunde0CHK %> onchange="submit();" /> Vis omsætning total <br /><input type="radio" name="FM_visprkunde" value="1" <%=visprkunde1CHK %> onchange="submit();" /> Vis omsætning pr. kunde / lev.
            
            &nbsp;top: 
            <select name="FM_toplist">
                <option value="5" <%=topSel5 %>>5</option>
                <option value="10" <%=topSel10 %>>10</option>
                <option value="20" <%=topSel20 %>>20</option>
                
            </select> 
            &nbsp;Vis for:
            <select name="FM_toplist_kunder_lev">
                <option value="0" <%=toplist_kunder_lev0SEL %>>Kunder</option>
                <option value="1" <%=toplist_kunder_lev1SEL %>>Leverandører</option>
               
            </select> 

            
            <span style="color:#999999;"><br />Fordelt efter faktureret omsætning, alle job uanset status.</span>

	</td>
	<td align=right>
	&nbsp;
	<%if request("print") <> "j" then%>
        <br />
	<input type="submit" value=" Kør >> ">
	<%end if%>
	
	</td>
	</tr>
	
	</form>
	</table>
	
	
	<!--filter div-->
	</td></tr>
	</table>
	</div>
	<%else %>

        <%if media <> "export" then %>

	    År: <%=strYear %>

        <%end if %>
	
   <%end if %>
	
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
<%
	
	
		'*****************************************************************************************************
		'**** Job og aftaler der skal indgå i omsætning, budget og forbrugstal *******************************
		'*****************************************************************************************************
		
		
		'*** valgte job ***
		call valgtejob()
		
		
		'**** Valgte aftaler *****
		call valgteaftaler()
		
		
		
		'*** For at spare (trimme) på SQL hvis der vælges alle job alle kunder og vis kun for jobanssvarlige ikke er slået til ****
		'*** Og der ikke er søgt på jobnavn ***
		if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
				
			'jobidFakSQLkri = " OR jobid <> 0 " HER 20150626 --> Hvis der er valgt job-status. Problemet optræder i oms. under fakturererede job
			jobnrSQLkri = " OR tjobnr <> '0' "
			jidSQLkri =  " OR id <> 0 "
			'seridFakSQLkri = " OR aftaleid <> 0 "
				
		end if
	
	
		'**************** Trimmer SQL states ************************
		len_jobidFakSQLkri = len(jobidFakSQLkri)
		right_jobidFakSQLkri = right(jobidFakSQLkri, len_jobidFakSQLkri - 3)
		jobidFakSQLkri =  right_jobidFakSQLkri
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		
		len_seridFakSQLkri = len(seridFakSQLkri)
		right_seridFakSQLkri = right(seridFakSQLkri, len_seridFakSQLkri - 3)
		seridFakSQLkri =  right_seridFakSQLkri
		'*****************************************************************************************************
	
	
	


	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	

    antalK = 0
    if cint(visprkunde) = 1 then
        

        redim fordeltPrkundeIds(20), fordeltPrkundeNames(20)
        if cint(kundeid) = 0 then 'top 20 kunder, faktureret / eller timeomsætning

                
                    if cint(toplist_kunder_lev) = 1 then 'Vis for leve
                    ktypeKri = " AND useasfak = 6"
                    else 'Vis for kunder : default
                    ktypeKri = " AND (useasfak = 0 OR useasfak = 1)"
                    end if

                    
                    
                    '("& jobidFakSQLkri &") AND
                    strSQLf = "SELECT f.fakadr, IF(faktype = 0, COALESCE(sum(f.beloeb * (f.kurs/100)),0), COALESCE(sum(f.beloeb * -1 * (f.kurs/100)),0)) AS beloeb, kkundenavn, kkundenr FROM fakturaer AS f "_
                    &" LEFT JOIN kunder AS k ON (kid = fakadr "& ktypeKri &") "_
                    &" WHERE "_
    	            &" ((YEAR(fakdato) = '"& strYear &"' AND brugfakdatolabel = 0) "_
                    &" OR (brugfakdatolabel = 1 AND YEAR(labeldato) = '"& strYear &"')) AND "_
					&" (faktype = 0) AND shadowcopy <> 1 "& ktypeKri &" GROUP BY fakadr ORDER BY beloeb DESC LIMIT "& topList 
        
                    'OR faktype = 1
                
                    'if session("mid") = 1 then
                    'response.write strSQLf
                    'response.flush 
                    'end if

                    oRec.open strSQLf, oConn, 3
                    while not oRec.EOF 
                    
                    'Response.write oRec("kkundenavn") & " - beløb: " & oRec("beloeb") & "<br>"

                    fordeltPrkundeIds(antalK) = oRec("fakadr")

                    'response.write fordeltPrkundeIds(antalK) & "<br>"

                    fordeltPrkundeNames(antalK) = oRec("kkundenavn")
                

                    'lastK = oRec("fakadr")
                    antalK = antalK + 1
                    oRec.movenext
                    wend
                    oRec.close


                    antalK = antalK - 1

        else

        antalK = 0

        end if

    else
    antalK = 0
    end if
    
    
    'response.end


    ekspTxt = ""

for k = 0 TO antalK 'antal kunder

	
	'*******************************************************************************************'
	'***** Finder registrerede timer og omsætning pr. job fordelt pr. måned. og medarbejere ****'
	'*******************************************************************************************'
	
     select case lto
     case "nt", "epi", "epi_no", "epi_uk", "xintranet - local"

             timerTotFakbar = 0
	        timerTotNotFakbar = 0
            mrdOmsTotal = 0
	        kostprisTot = 0


	
	        for m = 1 to 12

            mrdOms(m) = 0

            next

    
     case else
    
    
    redim mrdTotTimer(12)
	
	'** Key account: vis for alle medarb ***
	if cint(visKundejobans) = 1 then
	medarbSQlKri = " tmnr <> 0 "
	else
	medarbSQlKri = medarbSQlKri
	end if
	
	antalJobtimereg = 0
	strJobnrFundet = "#"
	
	
	call akttyper2009(2)

      
      if cint(visprkunde) = 1 then
      visprkundeSQLTimerKri = " AND tknr = " & fordeltPrkundeIds(k) 
      else
      visprkundeSQLTimerKri = ""
      end if


    
    timerTotFakbar = 0
	timerTotNotFakbar = 0
    mrdOmsTotal = 0
	kostprisTot = 0


	
	for m = 1 to 12

    mrdOms(m) = 0
	

   

	strSQL = "SELECT Tid, Tdato, t.timer, Tfaktim, Tjobnr, Tjobnavn, Taar, "_
	&" TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris, kostpris, t.fastpris, "_
	&" j.jobtpris, j.budgettimer, j.jobnavn, j.jobnr, t.kurs, j.jo_bruttooms "_
	&" FROM timer t "_
	&" LEFT JOIN job j ON (j.jobnr = tjobnr) "_
	&" WHERE ("& jobnrSQLkri &") AND "& replace(medarbSQlKri, "m.mid", "t.tmnr") &" AND ("& aty_sql_realhours &") AND "_
	&" YEAR(tdato) = "& strYear &" AND MONTH(tdato) = "& m &" " & visprkundeSQLTimerKri &" ORDER BY tid"' jobMedarbKri &" Tid > 0 ORDER BY Tid"
	
	'if m = 1 AND session("mid") = 1 then
	'Response.write strSQL &"<br><br>"
	'Response.flush
	'end if
	
	oRec.Open strSQL, oConn, 3
	while not oRec.EOF
					
							if instr(strJobnrFundet, "#"&oRec("Tjobnr")&"#") = 0 then
							strJobNavnogNr = strJobNavnogNr & oRec("jobnavn") &" ("&oRec("jobnr")&"), " 
							strJobnrFundet = strJobnrFundet & ",#"&oRec("Tjobnr")&"#"
							antalJobtimereg = antalJobtimereg + 1 
							end if
							
                           
							intJobTpris = oRec("jo_bruttooms") 'oRec("jobtpris")
                            intBudgetTimer = oRec("budgettimer")
						
	                        call akttyper2009Prop(oRec("tfaktim"))
							
							
							
							call omsBeregningprmd(m)
	
	
	
	
	'totTimer = totTimer + oRec("timer")
	oRec.movenext
	wend
    oRec.close

    
	
    
    
    next
	
	
	
	timerTotFakbar = mrdFakTimer(1) + mrdFakTimer(2) + mrdFakTimer(3) + mrdFakTimer(4) + mrdFakTimer(5) + mrdFakTimer(6) + mrdFakTimer(7) + mrdFakTimer(8) + mrdFakTimer(9) + mrdFakTimer(10) + mrdFakTimer(11) + mrdFakTimer(12)
	timerTotNotFakbar = mrdNotFakTimer(1) + mrdNotFakTimer(2) + mrdNotFakTimer(3) + mrdNotFakTimer(4) + mrdNotFakTimer(5) + mrdNotFakTimer(6) + mrdNotFakTimer(7) + mrdNotFakTimer(8) + mrdNotFakTimer(9) + mrdNotFakTimer(10) + mrdNotFakTimer(11) + mrdNotFakTimer(12)
	mrdOmsTotal = formatnumber(mrdOms(1) + mrdOms(2) + mrdOms(3) + mrdOms(4) + mrdOms(5) + mrdOms(6) + mrdOms(7) + mrdOms(8) + mrdOms(9) + mrdOms(10) + mrdOms(11) + mrdOms(12), 0)
	kostprisTot = formatnumber(mrdKostpris(1) + mrdKostpris(2) + mrdKostpris(3) + mrdKostpris(4) + mrdKostpris(5) + mrdKostpris(6) + mrdKostpris(7) + mrdKostpris(8) + mrdKostpris(9) + mrdKostpris(10) + mrdKostpris(11) + mrdKostpris(12), 0) '* Altid basisValISO_f8**'
	



    end select 'lto

    '*****************************************************************************************

	


	redim faktimer(12), fakbeloeb(12), fakmedarbBeloeb(12)
	redim faktimerPrMed(12), fakOmsJob(12), mrdForbMat(12), mrdForbMatindkob(12), mrdJobUdgifter(12)
	redim fakmatBeloeb(12), fakaktBeloeb(12), fakaktKorsBeloeb(12), fakaktTimBeloeb(12), fakaktStkBeloeb(12), fakaktEnhBeloeb(12), fakaktNoneBeloeb(12), kreditBelob(12)
	
    redim fakbelobJob(12), fakbelobAft(12), internFak(12), fakBelobEksInt(12) 
    redim mrdOrderqty(12), mrdProfit(12), mrdDbproc(12), mrdDBprocCnt(12)

	'********************'
	'*** Fakturering ****'
	'********************'
	
	
	'** Key account: vis for alle medarb ***
	if cint(visKundejobans) = 1 then
	fakmedspecSQLkri = " AND (fms.mid <> 0)"
	else
	fakmedspecSQLkri = fakmedspecSQLkri
	end if
	
	if cint(visprkunde) = 1 then
    visprkundeSQLFakKri = " AND fakadr = " & fordeltPrkundeIds(k) 
    else
    visprkundeSQLFakKri = ""
    end if
	
	
	
	antalFak = 0
	for m = 1 to 12
	
    miswrt = 0


	              
					
                    job_uaft_FakSQLkri = jobidFakSQLkri
					
				
	                '**** Faktureret på job og aftaler ==> TOTAL faktureret KUN efter faktura SYtem dato ***'
                    '**** Excl. faktura_medregn ikke i oms indstilling (interne) *******'
					strSQL4 = "SELECT f.fakdato, f.fid, COALESCE(sum(f.timer),0) AS timer, "_
					&" COALESCE(sum(f.beloeb * (f.kurs/100)),0) AS beloeb, f.kurs, "_
					&" f.faktype, faknr, f.aftaleid, f.jobid, medregnikkeioms"_
					&" FROM fakturaer f "_
					&" WHERE ((("& job_uaft_FakSQLkri &") OR ("& seridFakSQLkri &")) AND "_
					&" ((YEAR(fakdato) = '"& strYear &"' AND MONTH(fakdato) = '"& m &"' AND brugfakdatolabel = 0) "_
                    &" OR (brugfakdatolabel = 1 AND YEAR(labeldato) = '"& strYear &"' AND MONTH(labeldato) = '"& m &"'))) AND "_
					&" (faktype = 0 OR faktype = 1) AND shadowcopy <> 1 "& visprkundeSQLFakKri &" GROUP BY fid" 'fid
					

                   'response.write strSQL4 & "<br><br>"
                   'response.Flush

                    f = 0
					oRec2.open strSQL4, oConn, 3
					while not oRec2.EOF 
					
					
				    select case oRec2("faktype")
					case 0 '*** Faktura
					

                   
                  
				    if cint(oRec2("medregnikkeioms")) = 1 OR cint(oRec2("medregnikkeioms")) = 2 then
                    internFak(m) = internFak(m) + oRec2("beloeb")

                    else

                    faktimer(m) = faktimer(m) + oRec2("timer")
                    fakbelobrykker = 0 
                    ' + fakbelobrykker
					          
					
					fakbeloeb(m) = fakbeloeb(m) + oRec2("beloeb")   
    
                      '** Job / aftaler **'
                    if oRec2("jobid") <> 0 then
                    fakbelobJob(m) = fakbelobJob(m) + oRec2("beloeb")
                    else
                    fakbelobAft(m) = fakbelobAft(m) + oRec2("beloeb")
                    end if         



                    end if

                    kreditBelob(m) = kreditBelob(m)

					case 2 '*** Rykker --> Ikke med i query
					
					case 1 '** kreditnota
					faktimer(m) = faktimer(m) - oRec2("timer")
					fakbeloeb(m) = fakbeloeb(m) - oRec2("beloeb")
					
                    'strFakturaer = strFakturaer & oRec2("faknr")&", "
					'antalFak = antalFak + 1

                    kreditBelob(m) = kreditBelob(m) + oRec2("beloeb")


                    '** Job / aftaler **'
                    'if oRec2("jobid") <> 0 then 'kreditnotaer bliver ikke indregnet i job og aftale beløb
                    'fakbelobJob(m) = fakbelobJob(m) '- oRec2("beloeb")
                    'else
                    'fakbelobAft(m) = fakbelobAft(m) '- oRec2("beloeb")
                    'end if

                    'if cint(oRec2("medregnikkeioms")) = 1 then 'kreditnota kan ikke være intern
                    'internFak(m) = internFak(m) '- oRec2("beloeb")
                    'end if

					end select


                    if lastM <> m OR m = 1 AND miswrt = 0 then
                    strFakturaer = strFakturaer & "<br><br><b>"& monthname(m) &"</b><br>"
                    miswrt = 1
                    end if
                    lastM = m

					strFakturaer = strFakturaer & oRec2("faknr")&", "
					antalFak = antalFak + 1
                    
                    

                    if cint(visprkunde) = 0 OR (cint(visprkunde) = 1 AND (lto = "nt" OR lto = "intranet - local" OR lto = "essens")) then


					
										
										'**** Medarbejder udsp. ***
										strSQL5 = "SELECT sum(fms.beloeb * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fmsbel, sum(fms.fak) AS fmstimer, fms.fakid FROM fak_med_spec fms "_
										&"WHERE (fms.fakid = "& oRec2("fid") &" "& fakmedspecSQLkri &") GROUP BY fms.fakid"
									    
									    'if m = 1 then
										'Response.write "fakmedspecSQLkri: "& fakmedspecSQLkri & " -- " & strSQL5 &"<br>"
										'response.flush
										'end if
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakmedarbBeloeb(m) = fakmedarbBeloeb(m) + oRec3("fmsbel")
										faktimerPrMed(m) = faktimerPrMed(m) + oRec3("fmstimer")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakmedarbBeloeb(m) = fakmedarbBeloeb(m) - oRec3("fmsbel")
										faktimerPrMed(m) = faktimerPrMed(m) - oRec3("fmstimer")  
										end select
										
										
										end if
										oRec3.close
										
										
										'**** Materialer udsp. ***
                                        select case lto
                                        case "nt", "intranet - local"
                                        strSQL5 = "SELECT sum(fms.matantal) AS fmsbel, "_
										&" fms.matfakid FROM fak_mat_spec fms "_
										&"WHERE (fms.matfakid = "& oRec2("fid") &") GROUP BY fms.matfakid"

                                        case else
										strSQL5 = "SELECT sum(fms.matbeloeb * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fmsbel, "_
										&" fms.matfakid FROM fak_mat_spec fms "_
										&"WHERE (fms.matfakid = "& oRec2("fid") &") GROUP BY fms.matfakid"
                                        end select

										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakmatBeloeb(m) = fakmatBeloeb(m) + oRec3("fmsbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakmatBeloeb(m) = fakmatBeloeb(m) - oRec3("fmsbel")
										end select
										
										
										end if
										oRec3.close
										



                         end if 'if cint(visprkunde) = 0 then OR LTO


                                    
                          if cint(visprkunde) = 0 then
    
    
        
										
										'**** Akt sum udsp. excl. kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang <> 3) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktBeloeb(m) = fakaktBeloeb(m) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktBeloeb(m) = fakaktBeloeb(m) - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close
										

                                        '**** Akt sum udsp. kun timer ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = 0) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktTimBeloeb(m) = fakaktTimBeloeb(m) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktTimBeloeb(m) = fakaktTimBeloeb(m) - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close


                                         '**** Akt sum udsp. kun Stk ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = 1) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktStkBeloeb(m) = fakaktStkBeloeb(m) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktStkBeloeb(m) = fakaktStkBeloeb(m) - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close


                                        
                                         '**** Akt sum udsp. kun Enh. ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = 2) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktEnhBeloeb(m) = fakaktEnhBeloeb(m) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktEnhBeloeb(m) = fakaktEnhBeloeb(m) - oRec3("fdbel")
										end select
										
                                        end if
										oRec3.close


                                         '**** Akt sum udsp. kun Ikke angive (none -1). ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = -1) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktNoneBeloeb(m) = fakaktNoneBeloeb(m) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktNoneBeloeb(m) = fakaktNoneBeloeb(m) - oRec3("fdbel")
										end select
										
                                        end if
										oRec3.close
										
										'**** Akt sum udsp. KUN kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&" WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = 3) GROUP BY fd.fakid"
										
										'Response.write strSQL5 &"<br>"
										'response.flush
										
										oRec3.open strSQL5, oConn, 3
										if not oRec3.EOF then
										
										select case oRec2("faktype")
										case 0 '*** Faktura
										fakaktKorsBeloeb(m) = fakaktKorsBeloeb(m) + oRec3("fdbel")
										case 2 '*** Rykker --> Ikke med i query
										
										case 1 '** kreditnota
										fakaktKorsBeloeb(m) = fakaktKorsBeloeb(m) - oRec3("fdbel")
										end select
										
										
										end if
										oRec3.close
										
					

                    end if 'if cint(visprkunde) = 1 then

					
					oRec2.movenext
					wend
					oRec2.close
					
					
	
	next
	
	
	
	fakTimerTotal = faktimer(1) + faktimer(2) + faktimer(3) + faktimer(4) + faktimer(5) + faktimer(6) + faktimer(7) + faktimer(8) + faktimer(9) + faktimer(10) + faktimer(11) + faktimer(12)
	fakOmsTotal = fakbeloeb(1) + fakbeloeb(2) + fakbeloeb(3) + fakbeloeb(4) + fakbeloeb(5) + fakbeloeb(6) + fakbeloeb(7) + fakbeloeb(8) + fakbeloeb(9) + fakbeloeb(10) + fakbeloeb(11) + fakbeloeb(12)
	
    'faktimerKreKreTotal = faktimerKre(1) + faktimerKre(2) + faktimerKre(3) + faktimerKre(4) + faktimerKre(5) + faktimerKre(6) + faktimerKre(7) + faktimerKre(8) + faktimerKre(9) + faktimerKre(10) + faktimerKre(11) + faktimerKre(12)
	'fakOmsKreTotal = fakbeloebKre(1) + fakbeloebKre(2) + fakbeloebKre(3) + fakbeloebKre(4) + fakbeloebKre(5) + fakbeloebKre(6) + fakbeloebKre(7) + fakbeloebKre(8) + fakbeloebKre(9) + fakbeloebKre(10) + fakbeloebKre(11) + fakbeloebKre(12)
	

    kreditTotal = kreditBelob(1) + kreditBelob(2) + kreditBelob(3) + kreditBelob(4) + kreditBelob(5) + kreditBelob(6) + kreditBelob(7) + kreditBelob(8) + kreditBelob(9) + kreditBelob(10) + kreditBelob(11) + kreditBelob(12)
	

    fakOmsJobTotal = fakbelobJob(1) + fakbelobJob(2) + fakbelobJob(3) + fakbelobJob(4) + fakbelobJob(5) + fakbelobJob(6) + fakbelobJob(7) + fakbelobJob(8) + fakbelobJob(9) + fakbelobJob(10) + fakbelobJob(11) + fakbelobJob(12)
	fakOmsAftTotal = fakbelobAft(1) + fakbelobAft(2) + fakbelobAft(3) + fakbelobAft(4) + fakbelobAft(5) + fakbelobAft(6) + fakbelobAft(7) + fakbelobAft(8) + fakbelobAft(9) + fakbelobAft(10) + fakbelobAft(11) + fakbelobAft(12)
	fakOmsInternTotal = internFak(1) + internFak(2) + internFak(3) + internFak(4) + internFak(5) + internFak(6) + internFak(7) + internFak(8) + internFak(9) + internFak(10) + internFak(11) + internFak(12)
	
    fakOmsEksklInternTotal = (fakOmsTotal) + (kreditTotal) '(fakOmsInternTotal)
    fakBelobEksInt(1) = (fakbeloeb(1)) + (kreditBelob(1) + internFak(1))
    fakBelobEksInt(2) = (fakbeloeb(2)) + (kreditBelob(2) + internFak(2))  '- (internFak(2)
    fakBelobEksInt(3) = (fakbeloeb(3)) + (kreditBelob(3) + internFak(3))  '(internFak(3) + kreditBelob(3)) 
    fakBelobEksInt(4) = (fakbeloeb(4)) + (kreditBelob(4) + internFak(4))  '(internFak(4) + kreditBelob(4)) 
    fakBelobEksInt(5) = (fakbeloeb(5)) + (kreditBelob(5) + internFak(5))  '(internFak(5) + kreditBelob(5)) 
    fakBelobEksInt(6) = (fakbeloeb(6)) + (kreditBelob(6) + internFak(6))  '(internFak(6) + kreditBelob(6)) 
    fakBelobEksInt(7) = (fakbeloeb(7)) + (kreditBelob(7) + internFak(7))  '(internFak(7) + kreditBelob(7)) 
    fakBelobEksInt(8) = (fakbeloeb(8)) + (kreditBelob(8) + internFak(8))  '(internFak(8) + kreditBelob(8)) 
    fakBelobEksInt(9) = (fakbeloeb(9)) + (kreditBelob(9) + internFak(9))  '(internFak(9) + kreditBelob(9)) 
    fakBelobEksInt(10) = (fakbeloeb(10)) + (kreditBelob(10) + internFak(10))   '(internFak(10) + kreditBelob(10)) 
    fakBelobEksInt(11) = (fakbeloeb(11)) + (kreditBelob(11) + internFak(11))   '(internFak(11) + kreditBelob(11))
    fakBelobEksInt(12) = (fakbeloeb(12)) + (kreditBelob(12) + internFak(12))   '(internFak(12) + kreditBelob(12))  
    
    
    
    fakprMedarbTot = fakmedarbBeloeb(1) + fakmedarbBeloeb(2) + fakmedarbBeloeb(3) + fakmedarbBeloeb(4) + fakmedarbBeloeb(5) + fakmedarbBeloeb(6) + fakmedarbBeloeb(7) + fakmedarbBeloeb(8) + fakmedarbBeloeb(9) + fakmedarbBeloeb(10) + fakmedarbBeloeb(11) + fakmedarbBeloeb(12)
	faktimerPrMedTotal = faktimerPrMed(1) + faktimerPrMed(2) + faktimerPrMed(3) + faktimerPrMed(4) + faktimerPrMed(5) + faktimerPrMed(6) + faktimerPrMed(7) + faktimerPrMed(8) + faktimerPrMed(9) + faktimerPrMed(10) + faktimerPrMed(11) + faktimerPrMed(12)
	'fakOmsJobTotal = fakOmsJob(1) + fakOmsJob(2) + fakOmsJob(3) +  fakOmsJob(4) + fakOmsJob(5) + fakOmsJob(6) + fakOmsJob(7) + fakOmsJob(8) + fakOmsJob(9) + fakOmsJob(10) + fakOmsJob(11) + fakOmsJob(12)
	
	fakmatTot = fakmatBeloeb(1) + fakmatBeloeb(2) + fakmatBeloeb(3) + fakmatBeloeb(4) + fakmatBeloeb(5) + fakmatBeloeb(6) + fakmatBeloeb(7) + fakmatBeloeb(8) + fakmatBeloeb(9) + fakmatBeloeb(10) + fakmatBeloeb(11) + fakmatBeloeb(12)
	fakaktTot = fakaktBeloeb(1) + fakaktBeloeb(2) + fakaktBeloeb(3) + fakaktBeloeb(4) + fakaktBeloeb(5) + fakaktBeloeb(6) + fakaktBeloeb(7) + fakaktBeloeb(8) + fakaktBeloeb(9) + fakaktBeloeb(10) + fakaktBeloeb(11) + fakaktBeloeb(12)
	fakaktTimTot = fakaktTimBeloeb(1) + fakaktTimBeloeb(2) + fakaktTimBeloeb(3) + fakaktTimBeloeb(4) + fakaktTimBeloeb(5) + fakaktTimBeloeb(6) + fakaktTimBeloeb(7) + fakaktTimBeloeb(8) + fakaktTimBeloeb(9) + fakaktTimBeloeb(10) + fakaktTimBeloeb(11) + fakaktTimBeloeb(12)
	fakaktStkTot = fakaktStkBeloeb(1) + fakaktStkBeloeb(2) + fakaktStkBeloeb(3) + fakaktStkBeloeb(4) + fakaktStkBeloeb(5) + fakaktStkBeloeb(6) + fakaktStkBeloeb(7) + fakaktStkBeloeb(8) + fakaktStkBeloeb(9) + fakaktStkBeloeb(10) + fakaktStkBeloeb(11) + fakaktStkBeloeb(12)
	fakaktEnhTot = fakaktEnhBeloeb(1) + fakaktEnhBeloeb(2) + fakaktEnhBeloeb(3) + fakaktEnhBeloeb(4) + fakaktEnhBeloeb(5) + fakaktEnhBeloeb(6) + fakaktEnhBeloeb(7) + fakaktEnhBeloeb(8) + fakaktEnhBeloeb(9) + fakaktEnhBeloeb(10) + fakaktEnhBeloeb(11) + fakaktEnhBeloeb(12)
	fakaktNoneTot = fakaktNoneBeloeb(1) + fakaktNoneBeloeb(2) + fakaktNoneBeloeb(3) + fakaktNoneBeloeb(4) + fakaktNoneBeloeb(5) + fakaktNoneBeloeb(6) + fakaktNoneBeloeb(7) + fakaktNoneBeloeb(8) + fakaktNoneBeloeb(9) + fakaktNoneBeloeb(10) + fakaktNoneBeloeb(11) + fakaktNoneBeloeb(12)
	
    fakaktKorsTot = fakaktKorsBeloeb(1) + fakaktKorsBeloeb(2) + fakaktKorsBeloeb(3) + fakaktKorsBeloeb(4) + fakaktKorsBeloeb(5) + fakaktKorsBeloeb(6) + fakaktKorsBeloeb(7) + fakaktKorsBeloeb(8) + fakaktKorsBeloeb(9) + fakaktKorsBeloeb(10) + fakaktKorsBeloeb(11) + fakaktKorsBeloeb(12)
	
	
	

     

		'******************************************************************************'	
		'**** Finder budget (værdi på job) afhængig af jobansvarlige 			   ****' 
		'******************************************************************************'
		antalJobBudget = 0
        mrdOmsBudgetTOT = 0
        mrdOmsUdgifterTOT = 0
        MrdOrderQtyTOT = 0
		
		jobIdSQLKri = replace(job_uaft_FakSQLkri, "jobid", "id")

        if cint(visprkunde) = 1 then
        visprkundeSQLjobKri = " AND jobknr = " & fordeltPrkundeIds(k) 
        else
        visprkundeSQLjobKri = ""
        end if
		
		for m = 1 to 12

            mrdJobUdgifter(m) = 0
			mrdOmsBudget(m) = 0
            mrdOrderqty(m) = 0
            mrdDBprocCnt(m) = 0
		    

             select case lto
             case "nt", "xintranet - local"
             strSQL1 = "SELECT jobtpris, jobnavn, jobnr, budgettimer, fastpris, jobstartdato, jobslutdato, fakturerbart, udgifter, jo_bruttooms, orderqty, shippedqty, "_
             &" sales_price_pc, cost_price_pc, tax_pc, freight_pc, cost_price_pc_valuta, sales_price_pc_valuta, jobstatus, jo_dbproc"_
			 &" FROM job AS j WHERE ("& jobIdSQLKri &") AND YEAR(dt_confb_etd) = "& strYear &" AND MONTH(dt_confb_etd) = "& m &" "& jobstKri &" "& visprkundeSQLjobKri &" ORDER BY jobnavn" 'tilbud
		     ' &" ((orderqty * sales_price_pc) - (orderqty * cost_price_pc + (orderqty * cost_price_pc * tax_pc/100) + (orderqty * freight_pc))) AS profit "_ AND jobstatus <> 3
             case else
             strSQL1 = "SELECT jobtpris, jobnavn, jobnr, budgettimer, fastpris, jobstartdato, jobslutdato, fakturerbart, udgifter, jo_bruttooms"_
			 &" FROM job AS j WHERE ("& jobIdSQLKri &") AND YEAR(jobstartdato) = "& strYear &" AND MONTH(jobstartdato) = "& m &" AND fakturerbart = 1 "& jobstKri &" "& visprkundeSQLjobKri &" ORDER BY jobnavn" 'tilbud
		
             end select
                
		
			'Response.write "Budget:<br>"
			'Response.write strSQL1 &"<br>"
			oRec.open strSQL1, oConn, 0, 1
			while not oRec.EOF 
			        
			       mrdJobUdgifter(m) = mrdJobUdgifter(m) + oRec("udgifter") '**** Altid angivet i basisValISO_f8**'
                   mrdOmsBudget(m) = mrdOmsBudget(m) + oRec("jo_bruttooms")  'jobtpris '*** Altid angivet i basisValISO_f8 **'

                   select case lto
                   case "nt", "xintranet - local"

                   if cint(oRec("jobstatus")) = 1 OR cint(oRec("jobstatus")) = 3 then 
                   mrdOrderqty(m) = mrdOrderqty(m) + oRec("orderqty")
                   mrdOrderqtyThis = oRec("orderqty")
                   else
                   mrdOrderqty(m) = mrdOrderqty(m) + oRec("shippedqty")
                   mrdOrderqtyThis = oRec("shippedqty")
                   end if



                         
                        valutaKursOR3(oRec("sales_price_pc_valuta"))
                        salgsprisKurs = dblKursOR3

                        valutaKursOR3(oRec("cost_price_pc_valuta"))
                        kostprisKurs = dblKursOR3

                        tax_pc_calc = 0
                        profit_pc = 0
                        tax_pc_calc = (oRec("cost_price_pc")/1 * (oRec("tax_pc")/100))

                     
                        profit_pc = formatnumber((mrdOrderqtyThis * oRec("sales_price_pc") * (salgsprisKurs/100)) - (mrdOrderqtyThis * (oRec("cost_price_pc") + tax_pc_calc + oRec("freight_pc")) * (kostprisKurs/100)) * basisValKursUse/100, 2) 
                        
                        
                     
                
                   'mrdDBproc(m) = mrdDBproc(m)/1 + oRec("jo_dbproc")/1
                   'mrdDBprocCnt(m) = mrdDBprocCnt(m) + 1

                   mrdProfit(m) = mrdProfit(m)/1 + profit_pc/1

                  
                   end select

                
                   if lastM <> m then
                   strJob = strJob & "<br><br><b>"& monthname(m) &"</b><br>"
                   end if

                   strJob = strJob & oRec("jobnavn") & " ("& oRec("jobnr") &"), "
	                		
            lastM = m
			antalJobBudget = antalJobBudget + 1					
			oRec.movenext
			wend
			oRec.close

            'if mrdDBprocCnt(m) <> 0 then
            'mrdDBprocCnt(m) = mrdDBprocCnt(m)
            'else
            'mrdDBprocCnt(m) = 1
            'end if
            
            if mrdOmsBudget(m) <> 0 then
            mrdDBproc(m) = ((mrdProfit(m)*1/mrdOmsBudget(m)*1) * 100) 
            else
            mrdDBproc(m) = 0
            end if

             mrdDBproc(m) = formatnumber(mrdDBproc(m)/1, 2)
            'mrdDBproc(m) = formatnumber(mrdDBproc(m)/1 / mrdDBprocCnt(m)/1, 2)
	        'mrdDBproc(m) = mrdDBproc(m) * 1 	

		next	
		
		mrdOmsBudgetTOT = mrdOmsBudget(1) + mrdOmsBudget(2) + mrdOmsBudget(3) + mrdOmsBudget(4) + mrdOmsBudget(5) + mrdOmsBudget(6) + mrdOmsBudget(7) + mrdOmsBudget(8) + mrdOmsBudget(9) + mrdOmsBudget(10) + mrdOmsBudget(11) + mrdOmsBudget(12)
		mrdOmsUdgifterTOT = mrdJobUdgifter(1) + mrdJobUdgifter(2) + mrdJobUdgifter(3) + mrdJobUdgifter(4) + mrdJobUdgifter(5) + mrdJobUdgifter(6) + mrdJobUdgifter(7) + mrdJobUdgifter(8) + mrdJobUdgifter(9) + mrdJobUdgifter(10) + mrdJobUdgifter(11) + mrdJobUdgifter(12)
		MrdOrderQtyTOT = mrdOrderqty(1) + mrdOrderqty(2) + mrdOrderqty(3) + mrdOrderqty(4) + mrdOrderqty(5) +mrdOrderqty(6) + mrdOrderqty(7) + mrdOrderqty(8) +mrdOrderqty(9) +mrdOrderqty(10) + mrdOrderqty(11) + mrdOrderqty(12)   
        MrdProfitTOT = mrdProfit(1) + mrdProfit(2) + mrdProfit(3) + mrdProfit(4) + mrdProfit(5) + mrdProfit(6) + mrdProfit(7) + mrdProfit(8) + mrdProfit(9) + mrdProfit(10) + mrdProfit(11) + mrdProfit(12)  
        MrdDBprocTOT = mrdDBproc(1) + mrdDBproc(2) + mrdDBproc(3) + mrdDBproc(4) + mrdDBproc(5) + mrdDBproc(6) + mrdDBproc(7) + mrdDBproc(8) + mrdDBproc(9) + mrdDBproc(10) + mrdDBproc(11) + mrdDBproc(12)  

        'MrdDBprocTOT = MrdDBprocTOT / 12 'md

        if mrdOmsBudgetTOT <> 0 then
        MrdDBprocTOT = ((MrdProfitTOT*1/mrdOmsBudgetTOT*1) * 100)
        else
        MrdDBprocTOT = 0
        end if

       

        if cint(visprkunde) = 0 then

		
		'******************************************************************************'	
		'**** Finder værdi på Aftaler 											****' 
		'******************************************************************************'
		
		seridBudgetSQLkri = replace(seridFakSQLkri, "aftaleid", "id")
		antalAftaler = 0
		for m = 1 to 12
		
			strSQLaft = "SELECT pris, stdato, enheder, id, navn, aftalenr"_
			&" FROM serviceaft WHERE ("& seridBudgetSQLkri &") AND YEAR(stdato) = "& strYear &" AND MONTH(stdato) = "& m 
			
			'Response.write strSQLaft &"<br>"
			
			oRec.open strSQLaft, oConn, 0, 1
			while not oRec.EOF 
			
					mrdAftBudget(m) = mrdAftBudget(m) + formatnumber(oRec("pris"), 0) '** ALtid DKK **'
					strAftaler = strAftaler & oRec("navn") & " ("& oRec("aftalenr") &"), "
					
			antalAftaler = antalAftaler + 1				
			oRec.movenext
			wend
			oRec.close
		
		next	
		
		mrdAftBudgetTOT = formatnumber(mrdAftBudget(1) + mrdAftBudget(2) + mrdAftBudget(3) + mrdAftBudget(4) + mrdAftBudget(5) + mrdAftBudget(6) + mrdAftBudget(7) + mrdAftBudget(8) + mrdAftBudget(9) + mrdAftBudget(10) + mrdAftBudget(11) + mrdAftBudget(12), 0)
		
		'***** Brutto fortj. forkalk ***'
		mrdBruttoForkalkTot = 0
		redim mrdBruttoForkalk(12)
		for b = 1 to 12
		mrdBruttoForkalk(b) = (mrdOmsBudget(b) - mrdJobUdgifter(b)) + mrdAftBudget(b)
		mrdBruttoForkalkTot = mrdBruttoForkalkTot + (mrdBruttoForkalk(b))
		next
		
		mrdBruttoForkalkTot =  formatnumber(mrdBruttoForkalkTot, 2)
		
		'******************************************************************************'	
		'**** Finder materiale forbrug / udlæg	Salgspris		                   ****' 
		'******************************************************************************'
		
		
		
		medarbMATSQlKri = replace(lCase(medarbSQlKri), "m.mid", "usrid")
        medarbMATSQlKri = replace(lCase(medarbMATSQlKri), "tmnr", "usrid")
		for m = 1 to 12
		    
			strSQLmat = "SELECT (matantal * (matsalgspris * (kurs/100))) AS matsalgspris, (matantal * (matkobspris * (kurs/100))) AS matkobspris, kurs FROM materiale_forbrug"_
			&" WHERE ("& jobidFakSQLkri &") AND "& medarbMATSQlKri &" AND YEAR(forbrugsdato) = "& strYear &" AND MONTH(forbrugsdato) = "& m &" ORDER BY id" 
		
			'Response.write "Matforbrug:<br>"
			'Response.write strSQLmat &"<br>"
			'Response.Flush
			oRec.open strSQLmat, oConn, 0, 1
			while not oRec.EOF 
			
					mrdForbMat(m) = mrdForbMat(m) + oRec("matsalgspris")
					mrdForbMatindkob(m) = mrdForbMatindkob(m) + oRec("matkobspris")
					
			oRec.movenext
			wend
			oRec.close
		
		next	
		
		mrdForbMatTOT = mrdForbMat(1) + mrdForbMat(2) + mrdForbMat(3) + mrdForbMat(4) + mrdForbMat(5) + mrdForbMat(6) + mrdForbMat(7) + mrdForbMat(8) + mrdForbMat(9) + mrdForbMat(10) + mrdForbMat(11) + mrdForbMat(12)
		mrdForbMatindkobTOT = mrdForbMatindkob(1) + mrdForbMatindkob(2) + mrdForbMatindkob(3) + mrdForbMatindkob(4) + mrdForbMatindkob(5) + mrdForbMatindkob(6) + mrdForbMatindkob(7) + mrdForbMatindkob(8) + mrdForbMatindkob(9) + mrdForbMatindkob(10) + mrdForbMatindkob(11) + mrdForbMatindkob(12)
		
		'*************
	    

        end if ' if cint(visprkunde) = 0 then



	    
	   if k = 0 then
             
    
                if media <> "export" then   %>
    
                     <br>
                    <br><b>Brutto omsætning, og kostpris.</b>
                    <br> Alle beløb er i HELE <%=basisValISO_f8 %>. excl. moms<br>


                   <%if cint(visprkunde) = 1 then %>
                   <span style="color:#999999;">Omsætning pr. kunde er sorteret efter fakturabeløb før kreditnotaer, men vist incl. kreditnotaer, fordelt på top 5,10 eller 20 kunder.</span><br />
    
                <% end if 



	            tTop = 0
	            tLeft = 0
	            tWdth = 1204
	
	
	            call tableDiv(tTop,tLeft,tWdth)

                end if
	

    end if

	%>





<%




if request("print") <> "j" AND media <> "export" then
bd = 0
cs = 1
else
cs = 0
bd = 1
end if


if media <> "export" then
%>

<table cellspacing="<%=cs%>" cellpadding="2" border="0" >


        <%if cint(visprkunde) = 0 then %>
        <% call luft("Forkalkuleret") %>

        <%else %>


        <tr><td colspan="20"><br />&nbsp;</td></tr>

        <% call luft(left(fordeltPrkundeNames(k), 20)) %>

        <%

        end if

    
   
     
end if 'media
            
if media <> "export" then%>

<tr bgcolor="#8caae6">
	<td style="width:150px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff">&nbsp;A1) Budget Bruttooms. Job:</td>
    <td valign="bottom"  align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=basisValISO_f8 &" "&formatnumber(mrdOmsBudgetTOT,0)%></td>
    
    <%for m = 1 to 12 %>
     <td  valign="bottom" align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=basisValISO_f8 &" "&formatnumber(mrdOmsBudget(m), 0)%></font></td>
    <%next %>
    
</tr>

<%select case lto
  case "nt", "intranet - local"
  %>

    <tr bgcolor="#8caae6">
	<td style="width:150px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff">&nbsp;A2) Antal stk.:</td>
    <td valign="bottom"  align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdOrderQtyTOT,0)%></td>
    
    <%for m = 1 to 12 %>
     <td  valign="bottom" align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdOrderqty(m), 0)%></font></td>
    <%next %>
    
</tr>


<tr bgcolor="#8caae6">
	<td style="width:150px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff">&nbsp;A3) Profit:</td>
    <td valign="bottom"  align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=basisValISO_f8 &" "& formatnumber(mrdProfitTOT, 0)%></td>
    
    <%for m = 1 to 12 %>
     <td  valign="bottom" align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=basisValISO_f8 &" "& formatnumber(mrdProfit(m), 0)%></font></td>
    <%next %>
    
</tr>

<tr bgcolor="#8caae6">
	<td style="width:150px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff">&nbsp;A4) DB %:</td>
    <td valign="bottom"  align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdDBprocTOT, 2)%> %</td>
    
    <%for m = 1 to 12 %>
     <td  valign="bottom" align="right" style="width:70px; padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdDBproc(m), 2)%> %</font></td>
    <%next %>
    
</tr>


  <%
  end select %>

<%
else     

        if cint(visprkunde) = 1 then
        ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
        end if
      
         ekspTxt = ekspTxt & "Budget Bruttooms. Job;"

       for m = 1 to 12
         ekspTxt = ekspTxt & formatnumber(mrdOmsBudget(m), 0) & ";"  
        next 

    ekspTxt = ekspTxt & "xx99123sy#z"

    select case lto
    case "nt", "intranet - local"

         if cint(visprkunde) = 1 then
        ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
        end if
     
         ekspTxt = ekspTxt & "Antal stk.;"

       for m = 1 to 12
         ekspTxt = ekspTxt & formatnumber(mrdorderqty(m), 0) & ";"  
        next 

    ekspTxt = ekspTxt & "xx99123sy#z"



       if cint(visprkunde) = 1 then
        ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
        end if
     
         ekspTxt = ekspTxt & "Profit;"

       for m = 1 to 12
         ekspTxt = ekspTxt & formatnumber(mrdprofit(m), 0) & ";"  
        next 

    ekspTxt = ekspTxt & "xx99123sy#z"


      if cint(visprkunde) = 1 then
        ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
        end if
     
         ekspTxt = ekspTxt & "DB %;"

       for m = 1 to 12
         ekspTxt = ekspTxt & formatnumber(mrdDBproc(m), 2) & ";"  
        next 

    ekspTxt = ekspTxt & "xx99123sy#z"

    end select

    
end if 'media
      
 
    
select case lto 
case "nt", "intranet - local"

case else
      
if cint(visprkunde) = 0 then
    
  if media <> "export" then%> 

<tr bgcolor="#ffff99">
	<td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;A2) Udgift./Underlev.:</td>
    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdOmsUdgifterTOT, 0)%></td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(1), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(2), 0)%></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(3), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(4), 0)%></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(5), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(6), 0)%></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(7), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(8), 0)%></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(9), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(10), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(11), 0)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdJobUdgifter(12), 0)%></td>  
</tr>



<tr bgcolor="#EFf3FF">
	<td style="border-bottom:<%=bd%>px #000000 solid;"><font color="#000000">&nbsp;A3) Værdi Aftaler:</td>
    <td align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#000000"><%=basisValISO_f8 &" "&formatnumber(mrdAftBudgetTOT,0)%></td>
     <%for m = 1 to 12 %>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdAftBudget(m), 0)%></td>
    <%next %>		
</tr>



        <%if level = 1 then %>

        <tr bgcolor="#ffdfdf">
	        <td   style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;F1) Bruttofortj. forkalk:</td>
            <td valign="bottom"   align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdBruttoForkalkTOT,0)%></td>
    
            <%for m = 1 to 12 %>
            <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdBruttoForkalk(m), 0)%></td>
            <%next %>	
        </tr>

        <%end if %>


<%

  

 end if 'media   
 end if 'visprkunde 

 end select 'lto
    
    
    
    
    
 if media <> "export" then%>


        <%select case lto
           case "nt", "epi", "epi_no", "epi_uk", "intranet - local"
    
           case else %>

        <%if cint(visprkunde) = 0 then %>
        <% call luft("Realiseret") %>
        <%end if %>


        <tr bgcolor="#8caae6">
	        <td  style="border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff">&nbsp;B1) Realiseret Oms. ~ *:</td>
            <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=basisValISO_f8 &" "&formatnumber(mrdOmsTotal, 0)%></td>
    
            <%for m = 1 to 12 %>
            <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=basisValISO_f8 &" "&formatnumber(mrdOms(m), 0)%></td>
            <%
            next 
        
              %>
        </tr>

        <%end select %>


        
        <%select case lto
           case "nt", "intranet - local"
    
           case else %>

                <%if cint(visprkunde) = 0 then %>

                <tr bgcolor="orange">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;C1) Mat./Udlæg (købspris)*:</td>
                    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdForbMatindkobTOT,0)%></td>
    
    
                     <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdForbMatindkob(m), 0)%></td>
                    <%next %>	
    	
                </tr>

                <tr bgcolor="orange">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;C2) Mat./Udlæg (salgspris)*:</td>
                    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdForbMatTOT,0)%></td>

    
                 <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdForbMat(m), 0)%></td>
                    <%next %>	
                </tr>


                <%end if

            end select



else 


      select case lto
           case "nt", "epi", "epi_no", "epi_uk", "intranet - local"
    
           case else

                  if level = 1 then 

                        if cint(visprkunde) = 1 then
                        ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
                        end if
                     
                     ekspTxt = ekspTxt & "Realiseret oms.;"

                     for m = 1 to 12
                     ekspTxt = ekspTxt & formatnumber(mrdOms(m), 0) & ";"  
                    next 

                    ekspTxt = ekspTxt & "xx99123sy#z"

                 end if


        end select


end if 'media



        if media <> "export" then 
    
    
             if cint(visprkunde) = 0 then %>
            <% call luft("Faktureret") %>
            <%end if %>

                <tr bgcolor="#6CAE1C">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1) Udfaktureret oms.:</td>
                    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakOmsTotal, 0)%></td>

                 <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(m), 0)%></td>
                    <%next %>
                 </tr>

        <%else

              if cint(visprkunde) = 1 then
              ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
              end if 

             ekspTxt = ekspTxt & "Udfaktureret oms.;"
    
             for m = 1 to 12
                 ekspTxt = ekspTxt & formatnumber(fakbeloeb(m), 0) & ";"  
                next 

            ekspTxt = ekspTxt & "xx99123sy#z"


    

    
        end if %>



    <%if cint(visprkunde) = 0 then
        
        
                 if media <> "export" then %>

               <tr bgcolor="#CCCCCC">
	            <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1a) Kreditnotaer</td>
                <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(kreditTotal*-1, 0)%></td>

              <%for m = 1 to 12 %>
                <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(kreditBelob(m)*-1, 0)%></td>
                <%next %>
             </tr>


             <tr bgcolor="#CCCCCC">
	            <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1b) Interne fakturaer</td>
                <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakOmsInternTotal, 0)%></td>

             <%for m = 1 to 12 %>
                <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(internFak(m), 0)%></td>
                <%next %>
             </tr>

              <tr bgcolor="#6CAE1C">
	            <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1c) Udfaktureret beløb ialt: (incl. interne og før kredtin.)</td>
                <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakOmsEksklInternTotal, 0)%></td>

             <%for m = 1 to 12 %>
                <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbelobEksInt(m), 0)%></td>
                <%next %>
             </tr>

              <tr bgcolor="#ccff99">
	            <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1d) På job</td>
                <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakOmsJobTotal, 0)%></td>

             <%for m = 1 to 12 %>
                <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbelobJob(m), 0)%></td>
                <%next %>
             </tr>

              <tr bgcolor="#ccff99">
	            <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1e) På aftaler</td>
                <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakOmsAftTotal, 0)%></td>

             <%for m = 1 to 12 %>
                <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbelobAft(m), 0)%></td>
                <%next %>
             </tr>


             <%

                end if 'media

    end if 




    select case lto
    case "nt", "intranet - local", "essens"


        if lto = "essens" then
         d3kolonneTxt = "Materialer & udlæg"
        else
         d3kolonneTxt = "Antal stk. faktureret"
        end if


      if media <> "export" then 

     %>
     <tr bgcolor="yellowgreen">
	    <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D3) <%=d3kolonneTxt %>:</td>
        <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(fakmatTot, 0)%></td>
    
        <%for m = 1 to 12 %>
        <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(fakmatBeloeb(m), 0)%></td>
        <%next %>
     </tr>

     <% end if

          if cint(visprkunde) = 1 then
            ekspTxt = ekspTxt & fordeltPrkundeNames(k) & ";"  
          end if 

            ekspTxt = ekspTxt & ""& d3kolonneTxt &";"
    
        for m = 1 to 12
         ekspTxt = ekspTxt & formatnumber(fakmatBeloeb(m), 0) & ";"  
        next 

    ekspTxt = ekspTxt & "xx99123sy#z"


     
    ' case else

    end select




    if cint(visprkunde) = 0 then
        
        
          if media <> "export" then 


   


          select case lto
          case "nt", "intranet - local"

         case else


      call luft("Faktureret udspec.</b> <span style=""font-size:9px;""> (kreditnotaer fratrukket)</span><b>") %>
 


                <tr bgcolor="#DCF5BD">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D2) Pr. medarb (alle akt. typer).*:</td>
                    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakprMedarbTot, 0)%></td>
	
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakmedarbBeloeb(m), 0)%></td>
                    <%next %>
    
                </tr>

                <tr bgcolor="yellowgreen">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D3) Materialer & udlæg:</td>
                    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakmatTot, 0)%></td>
    
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakmatBeloeb(m), 0)%></td>
                    <%next %>
    
                 </tr>

                <tr bgcolor="yellowgreen">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D4) På sum aktiviteter (timer/stk/enh./ia):</td>
                    <td valign="bottom"  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktTot, 0)%></td>
	                <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktBeloeb(m), 0)%></td>
                    <%next %>
                </tr>


                <tr bgcolor="#DCF5BD">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D4a) På sum aktiviteter (kun timer):</td>
                        <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktTimTot, 0)%></td>
	
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktTimBeloeb(m), 0)%></td>
                    <%next %>
                    </tr>

                <tr bgcolor="#DCF5BD">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D4b) På sum aktiviteter (kun stk.):</td>
                        <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktStkTot, 0)%></td>
	
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktStkBeloeb(m), 0)%></td>
                    <%next %>
                    </tr>

                <tr bgcolor="#DCF5BD">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D4c) På sum aktiviteter (kun enheder):</td>
                        <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktEnhTot, 0)%></td>
	
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktEnhBeloeb(m), 0)%></td>
                    <%next %>
                    </tr>

                <tr bgcolor="#DCF5BD">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D4d) På sum aktiviteter (kun <u>I</u>kke <u>a</u>ngivet):</td>
                        <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktNoneTot, 0)%></td>
	
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktNoneBeloeb(m), 0)%></td>
                    <%next %>
                    </tr>

                <tr bgcolor="yellowgreen">
	                <td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D5) Kørsels aktiviteter (km):</td>
                        <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktKorsTot, 0)%></td>
	
                    <%for m = 1 to 12 %>
                    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakaktKorsBeloeb(m), 0)%></td>
                    <%next %>
                    </tr>



<% end select 'lto

 end if 'media

 end if 'if cint(visprkunde) = 0
    
%>


<%
 
   select case lto
    case "nt", "intranet - local"
     
   case else   
    
    
 if cint(visprkunde) = 0 then
    
 if media <> "export" then %>

<%if level = 1 then %>

<% call luft("Bruttofortjeneste") %>


<tr bgcolor="#FFFF99">
	<td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;E1) Intern kostpris realiseret*:</td>
    <td valign="bottom" width="80" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(kostprisTot, 0)%></td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(1), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(2), 0)%></font></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(3), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(4), 0)%></font></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(5), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(6), 0)%></font></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(7), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(8), 0)%></font></td> 
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(9), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(10), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(11), 0)%></font></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(mrdKostpris(12), 0)%></font></td>  
</tr>


<tr bgcolor="#ffdfdf">
	<td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;F2) Bruttofortj. realiseret:</td>
    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakOmsTotal - (mrdOmsUdgifterTOT + kostprisTot + mrdForbMatindkobTOT), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(1) - (mrdJobUdgifter(1) + mrdKostpris(1) + mrdForbMatindkob(1)), 0)%></td>
    <td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(2) - (mrdJobUdgifter(2) + mrdKostpris(2) + mrdForbMatindkob(2)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(3) - (mrdJobUdgifter(3) + mrdKostpris(3) + mrdForbMatindkob(3)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(4) - (mrdJobUdgifter(4) + mrdKostpris(4) + mrdForbMatindkob(4)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(5) - (mrdJobUdgifter(5) + mrdKostpris(5) + mrdForbMatindkob(5)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(6) - (mrdJobUdgifter(6) + mrdKostpris(6) + mrdForbMatindkob(6)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(7) - (mrdJobUdgifter(7) + mrdKostpris(7) + mrdForbMatindkob(7)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(8) - (mrdJobUdgifter(8) + mrdKostpris(8) + mrdForbMatindkob(8)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(9) - (mrdJobUdgifter(9) + mrdKostpris(9) + mrdForbMatindkob(9)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(10) - (mrdJobUdgifter(10) + mrdKostpris(10) + mrdForbMatindkob(10)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(11) - (mrdJobUdgifter(11) + mrdKostpris(11) + mrdForbMatindkob(11)), 0)%></td>
	<td  valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=basisValISO_f8 &" "&formatnumber(fakbeloeb(12) - (mrdJobUdgifter(12) + mrdKostpris(12) + mrdForbMatindkob(12)), 0)%></td>
</tr>
<%end if %>



<tr>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
   	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">&nbsp</td>
</tr>

<tr bgcolor="#ffffff">
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"><b>Timer<br /> Realiseret / Faktureret</b></td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Total</td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
   	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
</tr>

<tr bgcolor="#8caae6">
	<td width=110 style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;<font color="#ffffff">G) Fakturerbare timer*:</td>
    <td valign="bottom"  align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(timerTotFakbar, 2)%></td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(1), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(2), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(3), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(4), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(5), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(6), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(7), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(8), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(9), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(10), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(11), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdFakTimer(12), 2)%></td>
</tr>
<tr bgcolor="#8caae6">
	<td style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;<font color="#ffffff">H) Ikke fak.bare timer*:</td>
    <td valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(timerTotNotFakbar, 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(1), 2)%></td>
    <td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(2), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(3), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(4), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(5), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(6), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(7), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(8), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(9), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(10), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(11), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><font color="#ffffff"><%=formatnumber(mrdNotFakTimer(12), 2)%></td>
</tr>

<tr bgcolor="yellowgreen">
	<td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;J) Fak. pr. medarb. (timer/stk/enheder) *:</td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(fakTimerPrMedTotal, 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(1), 2)%></td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(2), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(3), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(4), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(5), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(6), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(7), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(8), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(9), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(10), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(11), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(12), 2)%></td>
</tr>
<tr bgcolor="#FFFFe1">
	<td  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;K) Balance:</td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber( fakTimerPrMedTotal - timerTotFakbar, 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(1) - mrdFakTimer(1), 2)%></td>
    <td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(2) - mrdFakTimer(2), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(3) - mrdFakTimer(3), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(4) - mrdFakTimer(4), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(5) - mrdFakTimer(5), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(6) - mrdFakTimer(6), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(7) - mrdFakTimer(7), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(8) - mrdFakTimer(8), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(9) - mrdFakTimer(9), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(10) - mrdFakTimer(10), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(11) - mrdFakTimer(11), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2px; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(12) - mrdFakTimer(12), 2)%></td>
</tr>
<%
    
end if 'media
    
end if 'visprkunde
   
 
    
end select 'lto %>


<%if media <> "export" then %>
</table>
<%end if %>




<%next 'antal kunder 

        if k = 0 then
    %>
        <br /><br />
        <table>
    <tr><td colspan="12">Der blev ikke fundet nogen kunder / leverandører der matchede de valgte kriterier.</td></tr>
            </table>
    <%
        response.end
 end if


if media <> "export" then %>
</div><!-- tableDiv -->
<%end if %>




        <%
            
         '******************* Eksport **************************' 
                if media = "export" then


    
                    call TimeOutVersion()
    


                 

	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	                datointerval = request("datointerval")
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\oms.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\omsexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\omsexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\omsexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\omsexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "omsexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				               strOskrifter = ""
				
				                if cint(visprkunde) = 1 then
                                strOskrifter = strOskrifter & "Kunde;"
                                end if
				                strOskrifter = strOskrifter & "Type;Januar;Februar;Marts;April;Maj;Juni;Juli;August;September;Oktober;November;December;"
				               

				
				                objF.writeLine("Periode afgrænsning: "& strYear & vbcrlf)
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				
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
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if 'media%>     
            
            
            
     









<br />
*) Tal i disse kolonner er baseret på den i pkt. 1) valgte medarbejder(e).<br />
Hvis <b>Key account</b> er slået til, er tallene baseret på alle medarbejdere på de job, hvor den <b>valgte medarbejder er jobansvarlig / kundeansvarlig.</b>
<br><br>


    <%
	if request("print") <> "j" then
	
	ptop = 0
    pleft = 840
    pwdt = 150
            
    call eksportogprint(ptop, pleft, pwdt)

    printEksLnk = "FM_segment="&segment&"&viskunabnejob0="&viskunabnejob0 &"&viskunabnejob1="&viskunabnejob1 &"&viskunabnejob2="&viskunabnejob2&"&seomsfor="&strYear&"&FM_aftaler="&aftaleid&"&FM_job="&jobid&"&FM_jobans="&jobans&"&FM_kundeans="&kundeans&"&FM_visprkunde="&visprkunde&"&FM_toplist="&toplist&"&FM_kunde="&kundeid&"&FM_medarb_hidden="&thisMiduse&"&FM_medarb="&thisMiduse&"&FM_progrupper="&progrp&"&FM_kundejobans_ell_alle="&visKundejobans&"&FM_jobsog="&jobSogValPrint
	%>
	<tr>
    <td align=center>
	<a href="<%=thisfile%>.asp?print=j&media=print&<%=printEksLnk %>" target="_blank">&nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
	</td><td><a href="<%=thisfile%>.asp?print=j&media=print&<%=printEksLnk %>" target="_blank" class=vmenu>Print version</a></td>
	</tr>
    <tr>
    <td align=center>
	<a href="#" onclick="Javascript:window.open('<%=thisfile%>.asp?media=export&<%=printEksLnk %>', '', 'width=350,height=150,resizable=no,scrollbars=no')" class=vmenu>&nbsp;<img src="../ill/export1.png" border=0 alt="" /></a>
	</td><td><a href="#" onclick="Javascript:window.open('<%=thisfile%>.asp?media=export&<%=printEksLnk %>', '', 'width=350,height=150,resizable=no,scrollbars=no')" class=vmenu>.CSV fil eksport</a></td>
	</tr>
	</table>
	</div>
	<%else %>
	        <% 
        Response.Write("<script language=""JavaScript"">window.print();</script>")
        %>
        	
	<%end if%>








<%if print <> "j" then%>

  <!--pagehelp-->

                <%

                itop = -16'56
                ileft = 595
                iwdt = 120
                ihgt = 0
                ibtop = 40 
                ibleft = 150
                ibwdt = 700
                ibhgt = 1210
                iId = "pagehelp"
                ibId = "pagehelp_bread"
                call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
                %>
                       
                       <b>Definitioner for beregninger af omsætning og fakturering</b>
			                <br /><br />
			                
			                
<b>A1) Budget Brutto oms. Job</b><br>
Samlet budget på job, beregnet udfra startdato på job i de viste måneder.<br>
Kun job der <u>ikke</u> er del af en aftale er medregnet. Tilbud er IKKE medregnet.<br><br>

<%select case lto
case "nt"%>
<b>A2) Antal stk.</b><br>
Antal stk. der skal produceres.
<br><br>
<%case else%>
<b>A2) Udgifter / Under-leverandører.</b><br>
Udgifter angivet på job til underleverandører.
<br><br>
<%end select %>

<b>A3) Værdi aftaler</b><br>
Samlet værdi angivet på aftaler, beregnet udfra aftale startdato i de viste måneder.<br>


<!-- B -->
<br /><b>B1) Realiseret omsætning. excl. moms.</b><br>
Beregnet således:<br />
<u>Antal realiserede faktuererbare timer * Medarbejdertypens timepris.</u><br />
<br><br>



<!-- C -->
<b>C1) Materialer</b><br />
Sum af udlæg / materiale forbrug på de valgte job. (Købspris)
<br /><br />
<b>C2) Materialer</b><br />
Sum af udlæg / materiale forbrug på de valgte job. (Salgspris)
<br /><br />

<!-- D -->

<br><br>

<b>D1) I) Udfaktureret omsætning excl. moms / timer</b><br>
Total faktureret beløb/timer job og aftaler.<br>
Kreditnotaer og Interne fakturaer er fratrukket.<br><!-- Interne fakturaer er medregnet, med mindre andet er angivet (D1a) -->
<br>


<b>D2-D5) J) Faktureret på de forskellige enhedstyper.</b><br>
Summen af de fakturerede beløb/timer, på hhv. timer, enheder, stk., ikke angivet og km.<br />
Kreditnotaer er fratrukket.<br><br>


<%if level = 1 then %>

<!-- E -->
<b>E1) Intern Kostpris.</b><br>
<u>Antal timer (på valgte medarbejdere) * Medarbejdertypens kostpris</u>
<br><br>


<b>F1) Bruttofortjeneste. Forkalk.</b><br>
(A1 + A3) - (A2)
<br><br>

<b>F2) Bruttofortjeneste. Real.</b><br>
(D1) - (C1 + E1)
<br><br>

<%end if %>

<b>K) Balance.</b><br>
Antal fakturerede timer, pr. medarb. (J) - Antal realiserede fakturerbare timer. (G)
<br><br>

<b>G) H) Timeforbrug.</b><br>
Antal timer på de valgte medarbejdere og job.<br>
Alle akt. typer der tæller med i dagligt timeregnskab er med. Også de ikke fakturerbare så som ikkefakturerbar tid, salg, ferie, sygdom og afspadsering.<br /><br>

			                
			                
			                
			                <br><br>
		<b>Job, aftale og faktura grundlag for ovenstående oversigt.</b><br />
			                
		<div style="width:600px; padding:10px; background-color:#Eff3ff; border:1px #8cAAe6 solid;">
		
		
		<font class=lillesort>
		
		<b>A1, B1, C1, C2, E1) Job:</b> (<%=antalJobBudget%>)<br>
		<%=strJob%>
		<br><br>
		
		<b>A3) Aftaler:</b> (<%=antalAftaler%>)<br>
		<%=strAftaler%>
		<br><br>
		
		<b>B1, E1, G, H) Job med timeregistreringer</b> (<%=antalJobtimereg%>)<br>
		<%=strJobNavnogNr%>
		<br><br>
		
		<b>B2) Fakturaer på job (også de job der er en del af en aftale) - IKKE AFTALER og interne fakturaer:</b> (<%=antalFakOms%>)<br>
		<%=strFakturaerOms%>
		<br /><br />
		
		<b>B2, D1, D2) Fakturaer:</b> (antal: <%=antalFak%>, excl. interne)<br>
		<%=strFakturaer%>
		<br><br> &nbsp;
		</font>
		
		
		</div><br /><br />&nbsp;
                		
                		
		                </td>
	                </tr></table></div>
	                
	                
	



	<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>&nbsp;
<%end if 
%>
</div><!-- side div --><%
end if

%>

<!--#include file="../inc/regular/footer_inc.asp"-->

       
