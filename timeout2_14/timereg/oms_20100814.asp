<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/inc_oms.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
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
	
	    if left(request("FM_medarb"), 1) = 0 then
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
	
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
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
			     fakmedspecSQLkri = "fms.mid = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
			     fakmedspecSQLkri = fakmedspecSQLkri & " OR fms.mid = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			jobAnsSQLkri =  "AND ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "AND (" & jobAns2SQLkri &")"
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
	jobnrSQLkri = " OR tjobnr = -1 "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	
	if len(request("viskunabnejob")) <> 0 then
	viskunabnejob = request("viskunabnejob")
	    
	    if viskunabnejob = 0 then
	    jost0CHK = "CHECKED"
	    jost1CHK = ""
	    else
	    jost1CHK = "CHECKED"
	    jost0CHK = ""
	    end if
	    
	else
	    if len(trim(request.cookies("stat")("viskunabnejob"))) <> 0 then
	    viskunabnejob = request.cookies("stat")("viskunabnejob")
	    
	            
	            if viskunabnejob = 0 then
                jost0CHK = "CHECKED"
                jost1CHK = ""
                else
                jost1CHK = "CHECKED"
                jost0CHK = ""
                end if
	            
	            
	    else
	    jost0CHK = "CHECKED"
	    jost1CHK = ""
	    viskunabnejob = 0
	    end if
	end if
	
	
	
	
	response.cookies("stat")("viskunabnejob") = viskunabnejob
	
	
	
	'************ slut faste filter var **************		
		
	
	
	
	
	'Response.write jobnrSQLkri
	
	
	
	
	%>
	<script>
	function clearJobsog(){
	document.getElementById("FM_jobsog").value = ""
	}
	</script>
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	
	<%else%>
	<html>
	<head>
		<title>TimeOut</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<table cellspacing="0" cellpadding="0" border="0" width="400">
	<tr>
		<td bgcolor="#003399"><img src="../ill/logo_topbar_print.gif" alt="" border="0" width="680" height="43"></td>
	</tr>
	</table><br>
	<%end if%>
	
	
	<%
	thisfile = "oms"
	if request("print") <> "j" then
	
	ptop = 204
    pleft = 840
    pwdt = 150
            
    call eksportogprint(ptop, pleft, pwdt)
	%>
	<tr>
    <td align=center>
	<a href="<%=thisfile%>.asp?print=j&seomsfor=<%=strYear%>&FM_aftaler=<%=aftaleid%>&FM_job=<%=jobid%>&FM_jobans=<%=jobans%>&FM_kundeans=<%=kundeans%>&FM_kunde=<%=kundeid%>&FM_medarb=<%=selmedarb%>&FM_progrupper=<%=progrp%>&FM_kundejobans_ell_alle=<%=visKundejobans%>&FM_jobsog=<%=jobSogValPrint%>" target="_blank">&nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
	</td><td><a href="<%=thisfile%>.asp?print=j&seomsfor=<%=strYear%>&FM_aftaler=<%=aftaleid%>&FM_job=<%=jobid%>&FM_jobans=<%=jobans%>&FM_kundeans=<%=kundeans%>&FM_kunde=<%=kundeid%>&FM_medarb=<%=selmedarb%>&FM_progrupper=<%=progrp%>&FM_kundejobans_ell_alle=<%=visKundejobans%>&FM_jobsog=<%=jobSogValPrint%>" target="_blank" class=rmenu>Print version</a></td>
	</tr>
	</table>
	</div>
	<%else %>
	        <% 
        Response.Write("<script language=""JavaScript"">window.print();</script>")
        %>
        	
	<%end if%>
	<!--include file="inc/stat_submenu.asp"-->
	<%
	if request("print") <> "j" then
	pleft = 20
	ptop = 132
	'ptopgrafik = 348
	else
	pleft = 10
	ptop = 65
	'ptopgrafik = 90
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible;">
	
	
	<%
	
	oimg = "ikon_stat_oms_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Omsætning"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	if request("print") <> "j" then 
	
	call filterheader(0,0,800,pTxt)%>
    	    
	  
	    <table cellspacing=0 cellpadding=10 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="oms.asp?rdir=<%=rdir %>" method="post">
	    
	<%end if %>
	    <%call medkunderjob %>
	    
	    <%if request("print") <> "j" then %>
	    
	    </td>
	    </tr>
	   
	
	
	<tr>
		<td valign=top><br /><b>År:</b>&nbsp;
	<select name="seomsfor" id="seomsfor" style="width:85px;">
			
			<%
			useaar = 2001
			for x = 0 to 10
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
	</td>
	<td align=right>
	&nbsp;
	<%if request("print") <> "j" then%>
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
	År: <%=strYear %>
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
		if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0  then 
				
			jobidFakSQLkri = " OR jobid <> 0 "
			jobnrSQLkri = " OR tjobnr <> 0 "
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
	
	
	'*******************************************************************************************'
	'***** Finder registrerede timer og omsætning pr. job fordelt pr. måned. og medarbejere ****'
	'*******************************************************************************************'
	dim mrdTotTimer(12)
	
	'** Key account: vis for alle medarb ***
	if cint(visKundejobans) = 1 then
	medarbSQlKri = " tmnr <> 0 "
	else
	medarbSQlKri = medarbSQlKri
	end if
	
	antalJobtimereg = 0
	strJobnrFundet = "#"
	
	
	call akttyper2009(2)
	
	for m = 1 to 12
	
	strSQL = "SELECT Tid, Tdato, t.timer, Tfaktim, Tjobnr, Tjobnavn, Taar, "_
	&" TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris, kostpris, t.fastpris, "_
	&" j.jobtpris, j.budgettimer, j.jobnavn, j.jobnr, t.kurs "_
	&" FROM timer t "_
	&" LEFT JOIN job j ON (j.jobnr = tjobnr) "_
	&" WHERE ("& jobnrSQLkri &") AND "& replace(medarbSQlKri, "m.mid", "t.tmnr") &" AND ("& aty_sql_realhours &") AND "_
	&" YEAR(tdato) = "& strYear &" AND MONTH(tdato) = "& m &" ORDER BY tid"' jobMedarbKri &" Tid > 0 ORDER BY Tid"
	
	'if m = 1 then
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
							
							intJobTpris = oRec("jobtpris")
							intBudgetTimer = oRec("budgettimer")
						
	                        call akttyper2009Prop(oRec("tfaktim"))
							
							'if cint(aty_fakbar) = 1 then
							'aktNotFakbar = "y"
							'else
							'aktNotFakbar = "n"
							'end if  
							
							
							'Response.Write "tid: " & oRec("tid") & " kp: " & oRec("kostpris") & "<br>"
							
							
							call omsBeregningprmd(m)
	
	
	
	
	totTimer = totTimer + oRec("timer")
	oRec.movenext
	wend
	
	oRec.close
	next
	
	
	
	timerTotFakbar = mrdFakTimer(1) + mrdFakTimer(2) + mrdFakTimer(3) + mrdFakTimer(4) + mrdFakTimer(5) + mrdFakTimer(6) + mrdFakTimer(7) + mrdFakTimer(8) + mrdFakTimer(9) + mrdFakTimer(10) + mrdFakTimer(11) + mrdFakTimer(12)
	timerTotNotFakbar = mrdNotFakTimer(1) + mrdNotFakTimer(2) + mrdNotFakTimer(3) + mrdNotFakTimer(4) + mrdNotFakTimer(5) + mrdNotFakTimer(6) + mrdNotFakTimer(7) + mrdNotFakTimer(8) + mrdNotFakTimer(9) + mrdNotFakTimer(10) + mrdNotFakTimer(11) + mrdNotFakTimer(12)
	mrdOmsTotal = formatcurrency(mrdOms(1) + mrdOms(2) + mrdOms(3) + mrdOms(4) + mrdOms(5) + mrdOms(6) + mrdOms(7) + mrdOms(8) + mrdOms(9) + mrdOms(10) + mrdOms(11) + mrdOms(12), 0)
	kostprisTot = formatcurrency(mrdKostpris(1) + mrdKostpris(2) + mrdKostpris(3) + mrdKostpris(4) + mrdKostpris(5) + mrdKostpris(6) + mrdKostpris(7) + mrdKostpris(8) + mrdKostpris(9) + mrdKostpris(10) + mrdKostpris(11) + mrdKostpris(12), 0) '* Altid DKK **'
	
	
	dim faktimer(12), fakbeloeb(12), fakmedarbBeloeb(12)
	dim faktimerPrMed(12), fakOmsJob(12), mrdForbMat(12), mrdForbMatindkob(12), mrdJobUdgifter(12)
	dim fakmatBeloeb(12), fakaktBeloeb(12), fakaktKorsBeloeb(12)
	
	'********************'
	'*** Fakturering ****'
	'********************'
	
	
	'** Key account: vis for alle medarb ***
	if cint(visKundejobans) = 1 then
	fakmedspecSQLkri = " AND (fms.mid <> 0)"
	else
	fakmedspecSQLkri = fakmedspecSQLkri
	end if
	
	
    
    '*** Fjerner de job der er en delaf aftale til fakturering ***'
    jobidkri = replace(jobidFakSQLkri, "jobid", "j.id")
    strSQLjob = "SELECT j.id, j.serviceaft FROM job j WHERE "& jobidkri
    
    job_uaft_FakSQLkri = ""
    
    oRec.open strSQLjob, oConn, 3
    while not oRec.EOF 

    'Response.Write "jid: " & oRec("id") & " s: " & oRec("serviceaft") & "<br>"
    if oRec("serviceaft") = 0 then
    job_uaft_FakSQLkri = job_uaft_FakSQLkri & " jobid = "& oRec("id") & " OR"
    end if

    oRec.movenext
    wend 
    oRec.close
    '*****
    
    
    len_job_uaft_FakSQLkri = len(job_uaft_FakSQLkri)
	if len_job_uaft_FakSQLkri > 0 then
	left_job_uaft_FakSQLkri = left(job_uaft_FakSQLkri,len_job_uaft_FakSQLkri-3)
	job_uaft_FakSQLkri = left_job_uaft_FakSQLkri
	else
	job_uaft_FakSQLkri = "jobid = 0"
	end if
	
	
	
	
	antalFak = 0
	for m = 1 to 12
	                
	                '***** Omsætning faktureret på job total, dvs både dem der er med på aftale og dem der ikke er ***'
	                strSQLoms = "SELECT  f.fakdato, f.fid, sum(f.beloeb * (f.kurs/100)) AS beloeb, "_
					&" f.faktype, faknr"_
					&" FROM fakturaer f "_
					&" WHERE (("& jobidFakSQLkri &") AND "_
					&" (YEAR(fakdato) = "& strYear &" AND MONTH(fakdato) = "& m &")) AND "_
					&" (faktype = 0 OR faktype = 1) AND shadowcopy <> 1 GROUP BY fid"
					
					'if m = 1 then
					'Response.write "<br><br>"& strSQLoms &"<hr>"
					'response.flush
					'end if
					
					f = 0
					oRec2.open strSQLoms, oConn, 3
					while not oRec2.EOF
					        
					        
					select case oRec2("faktype")
					case 0 '*** Faktura
					
					
		            omsbelobrykker = 0 
		            '*** Rykker beløb ***
		            strSQLrykker = "SELECT sum(rykkerbelob)AS rykkerbelob FROM faktura_rykker WHERE fakid = "& oRec2("fid")
		            'Response.Write strSQLrykker
		            'Response.flush
		            oRec3.open strSQLrykker, oConn, 3
		            if not oRec3.EOF then
		            omsbelobrykker = oRec3("rykkerbelob")
		            end if
		            oRec3.close
		            
		            if len(trim((omsbelobrykker))) <> 0 then
		            omsbelobrykker = omsbelobrykker
		            else
		            omsbelobrykker = 0
		            end if
					
					fakOmsJob(m) = fakOmsJob(m) + oRec2("beloeb") + omsbelobrykker
		            antalFak = antalFak + 1
					
					
					case 2 '*** Rykker --> Ikke med i query
					
					case 1 '** kreditnota
					
					fakOmsJob(m) = fakOmsJob(m) - oRec2("beloeb")
				    end select 
					
					strFakturaerOms = strFakturaerOms & oRec2("faknr")&", "
					antalFakOms = antalFakOms + 1
					
					oRec2.movenext
					wend
					oRec2.close
					
					
				
	                '**** Faktureret på job u. aftale samt aftaler ==> TOTAL faktureret ***'
					strSQL4 = "SELECT f.fakdato, f.fid, sum(f.timer) AS timer, "_
					&" sum(f.beloeb * (f.kurs/100)) AS beloeb, f.kurs, "_
					&" f.faktype, faknr"_
					&" FROM fakturaer f "_
					&" WHERE ((("& job_uaft_FakSQLkri &") OR ("& seridFakSQLkri &")) AND "_
					&" (YEAR(fakdato) = "& strYear &" AND MONTH(fakdato) = "& m &")) AND "_
					&" (faktype = 0 OR faktype = 1) AND shadowcopy <> 1 GROUP BY fid"
					
					
		            if m = 1 then
					'Response.write "<br><br>"& strSQL4 &"<br>"
					'response.flush
					end if
					
					f = 0
					oRec2.open strSQL4, oConn, 3
					while not oRec2.EOF 
					
					
				    select case oRec2("faktype")
					case 0 '*** Faktura
					faktimer(m) = faktimer(m) + oRec2("timer")
					
					            fakbelobrykker = 0 
					            '*** Rykker beløb ***
					            strSQLrykker = "SELECT sum(rykkerbelob)AS rykkerbelob FROM faktura_rykker WHERE fakid = "& oRec2("fid")
					            'Response.Write strSQLrykker
					            'Response.flush
					            oRec3.open strSQLrykker, oConn, 3
					            if not oRec3.EOF then
					            fakbelobrykker = oRec3("rykkerbelob")
					            end if
					            oRec3.close
					            
					            if len(trim((fakbelobrykker))) <> 0 then
					            fakbelobrykker = fakbelobrykker
					            else
					            fakbelobrykker = 0
					            end if
					
					fakbeloeb(m) = fakbeloeb(m) + oRec2("beloeb") + fakbelobrykker
					strFakturaer = strFakturaer & oRec2("faknr")&", "
					antalFak = antalFak + 1
					
					case 2 '*** Rykker --> Ikke med i query
					
					case 1 '** kreditnota
					faktimer(m) = faktimer(m) - oRec2("timer")
					fakbeloeb(m) = fakbeloeb(m) - oRec2("beloeb")
					strFakturaer = strFakturaer & oRec2("faknr")&", "
					antalFak = antalFak + 1
					end select
					
										
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
										strSQL5 = "SELECT sum(fms.matbeloeb * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fmsbel, "_
										&" fms.matfakid FROM fak_mat_spec fms "_
										&"WHERE (fms.matfakid = "& oRec2("fid") &") GROUP BY fms.matfakid"
										
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
										
										
										'**** Akt sum udsp. KUN kørsel ***
										strSQL5 = "SELECT sum(fd.aktpris * ("& replace(oRec2("kurs"), ",", ".") &"/100)) AS fdbel, "_
										&" fd.fakid FROM faktura_det fd "_
										&"WHERE (fd.fakid = "& oRec2("fid") &" AND showonfak = 1 AND enhedsang = 3) GROUP BY fd.fakid"
										
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
										
					
					
					oRec2.movenext
					wend
					oRec2.close
					
					
	
	next
	
	
	
	fakTimerTotal = faktimer(1) + faktimer(2) + faktimer(3) + faktimer(4) + faktimer(5) + faktimer(6) + faktimer(7) + faktimer(8) + faktimer(9) + faktimer(10) + faktimer(11) + faktimer(12)
	fakOmsTotal = fakbeloeb(1) + fakbeloeb(2) + fakbeloeb(3) + fakbeloeb(4) + fakbeloeb(5) + fakbeloeb(6) + fakbeloeb(7) + fakbeloeb(8) + fakbeloeb(9) + fakbeloeb(10) + fakbeloeb(11) + fakbeloeb(12)
	fakprMedarbTot = fakmedarbBeloeb(1) + fakmedarbBeloeb(2) + fakmedarbBeloeb(3) + fakmedarbBeloeb(4) + fakmedarbBeloeb(5) + fakmedarbBeloeb(6) + fakmedarbBeloeb(7) + fakmedarbBeloeb(8) + fakmedarbBeloeb(9) + fakmedarbBeloeb(10) + fakmedarbBeloeb(11) + fakmedarbBeloeb(12)
	faktimerPrMedTotal = faktimerPrMed(1) + faktimerPrMed(2) + faktimerPrMed(3) + faktimerPrMed(4) + faktimerPrMed(5) + faktimerPrMed(6) + faktimerPrMed(7) + faktimerPrMed(8) + faktimerPrMed(9) + faktimerPrMed(10) + faktimerPrMed(11) + faktimerPrMed(12)
	fakOmsJobTotal = fakOmsJob(1) + fakOmsJob(2) + fakOmsJob(3) +  fakOmsJob(4) + fakOmsJob(5) + fakOmsJob(6) + fakOmsJob(7) + fakOmsJob(8) + fakOmsJob(9) + fakOmsJob(10) + fakOmsJob(11) + fakOmsJob(12)
	
	fakmatTot = fakmatBeloeb(1) + fakmatBeloeb(2) + fakmatBeloeb(3) + fakmatBeloeb(4) + fakmatBeloeb(5) + fakmatBeloeb(6) + fakmatBeloeb(7) + fakmatBeloeb(8) + fakmatBeloeb(9) + fakmatBeloeb(10) + fakmatBeloeb(11) + fakmatBeloeb(12)
	fakaktTot = fakaktBeloeb(1) + fakaktBeloeb(2) + fakaktBeloeb(3) + fakaktBeloeb(4) + fakaktBeloeb(5) + fakaktBeloeb(6) + fakaktBeloeb(7) + fakaktBeloeb(8) + fakaktBeloeb(9) + fakaktBeloeb(10) + fakaktBeloeb(11) + fakaktBeloeb(12)
	fakaktKorsTot = fakaktKorsBeloeb(1) + fakaktKorsBeloeb(2) + fakaktKorsBeloeb(3) + fakaktKorsBeloeb(4) + fakaktKorsBeloeb(5) + fakaktKorsBeloeb(6) + fakaktKorsBeloeb(7) + fakaktKorsBeloeb(8) + fakaktKorsBeloeb(9) + fakaktKorsBeloeb(10) + fakaktKorsBeloeb(11) + fakaktKorsBeloeb(12)
	
	
	

		'******************************************************************************'	
		'**** Finder budget (værdi på job) afhængig af jobansvarlige 			   ****' 
		'******************************************************************************'
		antalJobBudget = 0
		
		jobIdSQLKri = replace(job_uaft_FakSQLkri, "jobid", "id")
		
		for m = 1 to 12
		    
			strSQL1 = "SELECT jobtpris, jobnavn, jobnr, budgettimer, fastpris, jobstartdato, jobslutdato, fakturerbart, udgifter"_
			&" FROM job WHERE ("& jobIdSQLKri &") AND YEAR(jobstartdato) = "& strYear &" AND MONTH(jobstartdato) = "& m &" AND fakturerbart = 1 AND jobstatus <> 3 ORDER BY jobnavn" 'tilbud
		
			'Response.write "Budget:<br>"
			'Response.write strSQL1 &"<br>"
			oRec.open strSQL1, oConn, 0, 1
			while not oRec.EOF 
			        
			        mrdJobUdgifter(m) = mrdJobUdgifter(m) + oRec("udgifter") '**** Altid angivet i DKK **'
					mrdOmsBudget(m) = mrdOmsBudget(m) + oRec("jobtpris")  '*** Altid angivet i DKK **'
					strJob = strJob & oRec("jobnavn") & " ("& oRec("jobnr") &"), "
			
			antalJobBudget = antalJobBudget + 1					
			oRec.movenext
			wend
			oRec.close
		
		next	
		
		mrdOmsBudgetTOT = mrdOmsBudget(1) + mrdOmsBudget(2) + mrdOmsBudget(3) + mrdOmsBudget(4) + mrdOmsBudget(5) + mrdOmsBudget(6) + mrdOmsBudget(7) + mrdOmsBudget(8) + mrdOmsBudget(9) + mrdOmsBudget(10) + mrdOmsBudget(11) + mrdOmsBudget(12)
		mrdOmsUdgifterTOT = mrdJobUdgifter(1) + mrdJobUdgifter(2) + mrdJobUdgifter(3) + mrdJobUdgifter(4) + mrdJobUdgifter(5) + mrdJobUdgifter(6) + mrdJobUdgifter(7) + mrdJobUdgifter(8) + mrdJobUdgifter(9) + mrdJobUdgifter(10) + mrdJobUdgifter(11) + mrdJobUdgifter(12)
		
		
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
		
		mrdAftBudgetTOT = formatcurrency(mrdAftBudget(1) + mrdAftBudget(2) + mrdAftBudget(3) + mrdAftBudget(4) + mrdAftBudget(5) + mrdAftBudget(6) + mrdAftBudget(7) + mrdAftBudget(8) + mrdAftBudget(9) + mrdAftBudget(10) + mrdAftBudget(11) + mrdAftBudget(12), 0)
		
		'***** Brutto fortj. forkalk ***'
		mrdBruttoForkalkTot = 0
		dim mrdBruttoForkalk(12)
		for b = 1 to 12
		mrdBruttoForkalk(b) = (mrdOmsBudget(b) - mrdJobUdgifter(b)) + mrdAftBudget(b)
		mrdBruttoForkalkTot = mrdBruttoForkalkTot + (mrdBruttoForkalk(b))
		next
		
		mrdBruttoForkalkTot =  formatcurrency(mrdBruttoForkalkTot, 2)
		
		'******************************************************************************'	
		'**** Finder materiale forbrug / udlæg	Salgspris		                   ****' 
		'******************************************************************************'
		
		
		
		medarbMATSQlKri = replace(lCase(medarbSQlKri), "m.mid", "usrid")
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
	    
	    
	   
    %>
    
    <br>
<br><b>Brutto omsætning, og kostpris.</b>
<br> Alle beløb er i HELE DKK. excl. moms<br>
    
    
    <%				

	tTop = 0
	tLeft = 0
	tWdth = 1204
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>





<%if request("print") <> "j" then
bd = 0
cs = 1
else
cs = 0
bd = 1
end if
%>

<table cellspacing="<%=cs%>" cellpadding="0" border="0" >

<% call luft("Forkalkuleret") %>

<tr bgcolor="#8caae6">
	<td style="width:150px; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff">&nbsp;A1) Brutto oms. Job:</td>
    <td valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudgetTOT,0)%></td>
    <td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(1), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(2), 0)%></font></td> 
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(3), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(4), 0)%></font></td> 
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(5), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(6), 0)%></font></td> 
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(7), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(8), 0)%></font></td> 
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(9), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(10), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(11), 0)%></font></td>
	<td  valign="bottom"  align="right" style="width:70px; padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsBudget(12), 0)%></font></td>  
</tr>

<tr bgcolor="#ffff99">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;A2) Udgift./Underlev.:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdOmsUdgifterTOT, 0)%></td>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(1), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(2), 0)%></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(3), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(4), 0)%></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(5), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(6), 0)%></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(7), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(8), 0)%></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(9), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(10), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(11), 0)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdJobUdgifter(12), 0)%></td>  
</tr>
<tr bgcolor="#EFf3FF">
	<td style="border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#000000">&nbsp;A3) Værdi Aftaler:</td>
    <td align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#000000"><%=formatcurrency(mrdAftBudgetTOT,0)%></td>
     <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdAftBudget(m), 0)%></td>
    <%next %>		
</tr>


<%if level = 1 then %>

<tr bgcolor="#ffdfdf">
	<td class=lille  style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;F1) Bruttofortj. forkalk:</td>
    <td valign="bottom" class=lille  align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdBruttoForkalkTOT,0)%></td>
    
    <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdBruttoForkalk(m), 0)%></td>
    <%next %>	
</tr>

<%end if %>


<% call luft("Realiseret") %>


<tr bgcolor="#8caae6">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff">&nbsp;B1) Realiseret Oms. ~ *:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOmsTotal, 0)%></td>
    
    <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(mrdOms(m), 0)%></td>
    <%next %>
</tr>
<tr bgcolor="#8caae6">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff">&nbsp;B2) Faktureret Omsætning:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(fakOmsJobTotal, 0)%></td>
    </td>
	 <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatcurrency(fakOmsJob(m), 0)%></td>
    <%next %>	
	</tr>

<tr bgcolor="orange">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;C1) Mat./Udlæg (købspris)*:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdForbMatindkobTOT,0)%></td>
    
    
     <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdForbMatindkob(m), 0)%></td>
    <%next %>	
    	
</tr>

<tr bgcolor="orange">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;C2) Mat./Udlæg (salgspris)*:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdForbMatTOT,0)%></td>

    
 <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdForbMat(m), 0)%></td>
    <%next %>	
</tr>


<% call luft("Faktureret") %>

<tr bgcolor="#6CAE1C">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D1) Faktureret Beløb:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakOmsTotal, 0)%></td>

 <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(m), 0)%></td>
    <%next %>
 </tr>

<tr bgcolor="#DCF5BD">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D2) Pr. medarb (alle akt. typer).*:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakprMedarbTot, 0)%></td>
	
    <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakmedarbBeloeb(m), 0)%></td>
    <%next %>
    
</tr>

<tr bgcolor="yellowgreen">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D3) Materialer:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakmatTot, 0)%></td>
    
    <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakmatBeloeb(m), 0)%></td>
    <%next %>
    
 </tr>

<tr bgcolor="yellowgreen">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D4) På sum aktiviteter (timer/stk/enheder):</td>
    <td valign="bottom" class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakaktTot, 0)%></td>
	<%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakaktBeloeb(m), 0)%></td>
    <%next %>
</tr>

<tr bgcolor="yellowgreen">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;D5) Kørsels aktiviteter (km):</td>
        <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakaktKorsTot, 0)%></td>
	
    <%for m = 1 to 12 %>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakaktKorsBeloeb(m), 0)%></td>
    <%next %>
    </tr>

<%if level = 1 then %>

<% call luft("Bruttofortjeneste") %>


<tr bgcolor="#FFFF99">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;E1) Intern kostpris real.*:</td>
    <td valign="bottom" width="80"class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(kostprisTot, 0)%></td>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(1), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(2), 0)%></font></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(3), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(4), 0)%></font></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(5), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(6), 0)%></font></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(7), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(8), 0)%></font></td> 
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(9), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(10), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(11), 0)%></font></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(mrdKostpris(12), 0)%></font></td>  
</tr>


<tr bgcolor="#ffdfdf">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;F2) Bruttofortj. Real.:</td>
    <td valign="bottom" class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakOmsTotal - (mrdOmsUdgifterTOT + kostprisTot + mrdForbMatindkobTOT), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(1) - (mrdJobUdgifter(1) + mrdKostpris(1) + mrdForbMatindkob(1)), 0)%></td>
    <td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(2) - (mrdJobUdgifter(2) + mrdKostpris(2) + mrdForbMatindkob(2)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(3) - (mrdJobUdgifter(3) + mrdKostpris(3) + mrdForbMatindkob(3)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(4) - (mrdJobUdgifter(4) + mrdKostpris(4) + mrdForbMatindkob(4)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(5) - (mrdJobUdgifter(5) + mrdKostpris(5) + mrdForbMatindkob(5)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(6) - (mrdJobUdgifter(6) + mrdKostpris(6) + mrdForbMatindkob(6)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(7) - (mrdJobUdgifter(7) + mrdKostpris(7) + mrdForbMatindkob(7)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(8) - (mrdJobUdgifter(8) + mrdKostpris(8) + mrdForbMatindkob(8)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(9) - (mrdJobUdgifter(9) + mrdKostpris(9) + mrdForbMatindkob(9)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(10) - (mrdJobUdgifter(10) + mrdKostpris(10) + mrdForbMatindkob(10)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(11) - (mrdJobUdgifter(11) + mrdKostpris(11) + mrdForbMatindkob(11)), 0)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatcurrency(fakbeloeb(12) - (mrdJobUdgifter(12) + mrdKostpris(12) + mrdForbMatindkob(12)), 0)%></td>
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
	<td width=110 style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;<font size="1" color="#ffffff">G) Fakturerbare timer*:</td>
    <td valign="bottom"  align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(timerTotFakbar, 2)%></td>
    <td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(1), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(2), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(3), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(4), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(5), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(6), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(7), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(8), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(9), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(10), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(11), 2)%></td>
	<td  valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdFakTimer(12), 2)%></td>
</tr>
<tr bgcolor="#8caae6">
	<td style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;<font size="1" color="#ffffff">H) Ikke fak.bare timer*:</td>
    <td valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(timerTotNotFakbar, 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(1), 2)%></td>
    <td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(2), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(3), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(4), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(5), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(6), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(7), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(8), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(9), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(10), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(11), 2)%></td>
	<td valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><font size="1" color="#ffffff"><%=formatnumber(mrdNotFakTimer(12), 2)%></td>
</tr>
<!--
<tr bgcolor="yellowgreen">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;I) Faktureret:</td>
    <td class=lille align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(fakTimerTotal, 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(1), 2)%></td>
    <td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(2), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(3), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(4), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(5), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(6), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(7), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(8), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(9), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(10), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(11), 2)%></td>
	<td class=lille valign="bottom"align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimer(12), 2)%></td>
</tr>
-->
<tr bgcolor="yellowgreen">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;J) Fak. pr. medarb. (timer/stk/enheder) *:</td>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(fakTimerPrMedTotal, 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(1), 2)%></td>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(2), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(3), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(4), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(5), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(6), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(7), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(8), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(9), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(10), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(11), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(12), 2)%></td>
</tr>
<tr bgcolor="#FFFFe1">
	<td class=lille style="border-bottom:<%=bd%>px #000000 solid;">&nbsp;K) Balance:</td>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber( fakTimerPrMedTotal - timerTotFakbar, 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(1) - mrdFakTimer(1), 2)%></td>
    <td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(2) - mrdFakTimer(2), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(3) - mrdFakTimer(3), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(4) - mrdFakTimer(4), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(5) - mrdFakTimer(5), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(6) - mrdFakTimer(6), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(7) - mrdFakTimer(7), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(8) - mrdFakTimer(8), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(9) - mrdFakTimer(9), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(10) - mrdFakTimer(10), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(11) - mrdFakTimer(11), 2)%></td>
	<td class=lille valign="bottom" align="right" style="padding-right:2; border-bottom:<%=bd%>px #000000 solid;"><%=formatnumber(faktimerPrMed(12) - mrdFakTimer(12), 2)%></td>
</tr>
</table>




</div><!-- tableDiv -->

<br />
*) Tal i disse kolonner er baseret på den i pkt. 1) valgte medarbejder(e).<br />
Hvis <b>Key account</b> er slået til, er tallene baseret på alle medarbejdere på de job, hvor den <b>valgte medarbejder er jobansvarlig / kundeansvarlig.</b>
<br><br>



<%if print <> "j" then%>

  <!--pagehelp-->

                <%

                itop = 56
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
			                
			                
<b>A1) Brutto oms. Job</b><br>
Samlet værdi på job, beregnet udfra startdato på job i de viste måneder.<br>
Kun job der <u>ikke</u> er del af en aftale er medregnet.<br><br>

<b>A2) Udgifter / Under-leverandører.</b><br>
Udgifter angivet på job til underleverandører.
<br><br>

<b>A3) Værdi aftaler</b><br>
Samlet værdi angivet på aftaler, beregnet udfra aftale startdato i de viste måneder.<br>


<!-- B -->
<br /><b>B1) Realiseret omsætning. excl. moms.</b><br>
Beregnet således:<br />
Summen af lbn. timer og fastpris job for de valgte medarbejdere.<br>
For fastpris job gælder:<br>
<u>Antal registrerede faktuererbare timer * (Beløb på job / Forkalk. fakturerbare timer)</u><br>
For job, der er sat til løbende timer, gælder:<br>
<u>Antal reg. faktuererbare timer x Medarbejdertypens timepris.</u><br />
<br><br>


<b>B2) Faktureret omsætning. excl. moms.</b><br>
Summen af fakturerede beløb på ALLE job (IKKE AFTALER) uanset om jobbet er en del af en aftale. <br />
<b>Fakturaer</b> på job der er med på en aftale bliver ikke medregnet i D1-D5.
<br><br>
<!-- C -->
<b>C1) Materialer</b><br />
Sum af udlæg / materiale forbrug på de valgte job. (Købspris)
<br /><br />
<b>C2) Materialer</b><br />
Sum af udlæg / materiale forbrug på de valgte job. (Salgspris)
<br /><br />

<!-- D -->
<b>D1) I) Total faktureret beløb excl. moms / timer</b><br>
Total faktureret beløb/timer på job der ikke er en del af en aftale, plus aftale fakturaer.<br>
<b>Fakturaer oprettet op job der er en del af en aftale er ikke medregnet.</b><br />
Rykker-beløb indregnet, og kreditnotaer er fratrukket.<br>
<br>


<b>D2) J) Faktureret på de valgte medarbejdere.</b><br>
Summen af de fakturerede beløb/timer, på de(n) <u>valgte medarbejdere</u> (for hver enkelt medarbejderlinie på hver aktivitet). 
På fakturaer oprettet på job der ikke er en del af en aftale, samt på aftaler.
Rykker-skrivelser er ikke indregnet, men kreditnotaer er fratrukket.<br><br>


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
		
		<b>B2) Fakturaer på alle job - IKKE AFTALER - også de job der er en del af en aftale..:</b> (<%=antalFakOms%>)<br>
		<%=strFakturaerOms%>
		<br /><br />
		
		<b>B2, D1, D2) Fakturaer:</b> (<%=antalFak%>)<br>
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
