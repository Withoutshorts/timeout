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
	
    menu = request("menu")
    thisfile = "pipeline"

    media = request("media")
    if media = "print" OR media = "export" then
        print = "j"
    end if

	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
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
	
    if len(trim(request("FM_start_mrd"))) <> 0 then
	strMrd = request("FM_start_mrd")
	else
	strMrd = month(now)
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
	'if len(request("FM_aftaler")) <> 0 then ' AND jobid <= 0 AND len(jobSogVal) = 0 then
	'aftaleid = request("FM_aftaler")
		
		'if cint(aftaleid) > 0 then 
		'jobid = 0
		'end if
		
	'else
		'if cint(jobid) > 0 OR len(jobSogVal) <> 0  then
		aftaleid = -1
		'else
	'	aftaleid = 0
		'end if
	'end if
	
	
	
	'**** Alle SQL kri starter med NUL records ****
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	
    
       
	   
	
	
	
	
	
	
	'************ slut faste filter var **************		
	

	if media = "print" OR media = "export" then
	'leftPos = 20
	'topPos = 62
    %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	else
	'leftPos = 90
	'topPos = 102
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

  
    <script src="inc/pipeline_jav.js"></script>
	<!--include file="../inc/regular/topmenu_inc.asp"-->

	
	<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%
        call tsamainmenu(7)
    %>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
        call stattopmenu()
    %>
	</div>
        -->



	<%

        call menu_2014()

	end if
	%>


    <!--#include file="inc/dato2_b.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	
	<%if media <> "print" AND media <> "export" then%>
	
	<%
    'oleftpx = 1065
	'otoppx = 115
	'owdtpx = 120
	
	'call opretNy("obs.asp?menu=job&func=opret&id=0&int=1&rdir=pipe", "Opret nyt job", otoppx, oleftpx, owdtpx) 
	
    %>

	
	
	
    
    
    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:220px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%
	
	exp_loadtid = 0
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	ca.: <b>3-5 sek.</b>
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>
	
    <%end if%>

	
	<%
	if media <> "print" AND media <> "export" then
	pleft = 90
	ptop = 102 '132
	'ptopgrafik = 348
	else
	pleft = 10
	ptop = 20
	'ptopgrafik = 90
	end if

    if media <> "export" then
	%>	
	<div id="sidediv" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
	
	
	<%
	
	oimg = "ikon_gantt_48.png"
	oleft = 0
	otop = 0
	owdt = 200
	oskrift = "Pipeline"
	
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
    end if
	
	if media <> "print" AND media <> "export" then 
	
	call filterheader_2013(0,0,800,oskrift)%>
    	    
	  
	    <table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="pipeline.asp?rdir=<%=rdir %>" method="post">
	    
	<%end if %>
	    
        
        <%call medkunderjob %>
	    
	    <%if media <> "print" AND media <> "export" then %>
	    
	    </td>
	    </tr>
	   
	
	
	<tr>
		<td valign=top>
        <br />
        <b>Fra:</b> 
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

     og 12 md frem.

	</td>
	<td align=right>
	&nbsp;<input type="submit" value=" Kør >> ">
	</td></tr>
    </form>
	</table>
	
	
	<!--filter div-->
	</td></tr>
	</table>
	</div>
	
	<%end if %>
	
	
	
	
	
        <%
	    '*****************************************************************************************************
		'**** Job og aftaler der skal indgå i omsætning, budget og forbrugstal *******************************
		'*****************************************************************************************************
		
		
		'*** valgte job ***
		call valgtejob()
		
        'Response.write "jobnrSQLkri" & jobnrSQLkri

		
		'**** Valgte aftaler *****
		'call valgteaftaler()
		
		
		
		'*** For at spare (trimme) på SQL hvis der vælges alle job alle kunder og vis kun for jobanssvarlige ikke er slået til ****
		'*** Og der ikke er søgt på jobnavn ***
		if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
				
			'jobidFakSQLkri = " OR jobid <> 0 "
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
	
	
	

    if media <> "export" then
	
    tTop = 20
	tLeft = 0
	tWdth = 1300
	
	
	call tableDiv(tTop,tLeft,tWdth)

    end if

    
    dim strMrdno(11)
    dim strMrdnm(11)
    dim strYearno(11)
    dim strDayno(11)
    dim tableHTML(11)
    dim mdIdTot(11)
    for m = 0 to 11

    strMrdno(m) = strMrd + m
    if strMrdno(m) <= 12 then
    strMrdno(m) = strMrdno(m)
    strYearno(m) = strYear
    else
    strMrdno(m) = strMrdno(m) - 12
    strYearno(m) = strYear + 1
    end if

    if m = 11 then
    strMrdEnd = strMrdno(m)
    strYearEnd = strYearno(m) 
    end if

    strMrdnm(m) = monthname(strMrdno(m))


    select case strMrdno(m)
    case 1,3,5,7,8,10,12
    strDayno(m) = 31
    case 2
        select case strYearno(m)
        case 2000, 2004,2008,2012,2016,2020,2024,2028,2032,2036,2040,2044
        strDayno(m) = 29
        case else
        strDayno(m) = 28
        end select
    case else
    strDayno(m) = 30
    end select


    if m = 11 then
    strDayEnd = strDayno(m)
    end if

    tableHTML(m) = "<table cellspacing=1 cellpadding=0 border=0 width='100%'>"

    next
    


   

    if media <> "export" then %>
	
    <table cellspacing="0" cellpadding="0" border="0"><tr><td valign=top style="width:504px;">
	
	<table cellspacing="1" cellpadding="0" border="0" bgcolor="#d6dff5" width=100%>
	<tr bgcolor="#5582D2">
    <%
    for m = 0 to 11
    %>
    <td class=alt style="width:40px; padding:2px;" valign=bottom class=lille><%=left(strMrdnm(m), 3) &" "& right(strYearno(m), 2) %></td>
    <%
    next
     %></tr><%

    


	'strBudget = "<table cellspacing=1 cellpadding=0 border=1>"
	strJob = "<table cellspacing=0 cellpadding=0 border=0><tr bgcolor=""#5582d2""><td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid; width:250px;""><b>Jobnavn (nr)</b><br>Kunde (knr)</td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Brutto oms. "&basisValISO&"</b>"_
    &"<td class=alt valign=bottom style=""padding:2px;"" class=alt><b>Startdato</b> (måneder)</td>"_
    &"<td class=alt valign=bottom style=""padding:2px;""><b>Sand. %</b></td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;""><b>Pipeline værdi</b><br>(Brutto Oms - Udgifter ulev. * Sand.) "& basisValISO &"</td>"_
    &"<td style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom class=lille colspan=3><b>Job-ansv.</b><br>Værdi "& basisValISO &"</td>"_
    &"<td style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom class=lille colspan=3><b>Job-ejer</b><br>Værdi "& basisValISO &"</td>"_
    &"<td style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom class=lille colspan=3><b>Med-ansv.1</b><br>Værdi "& basisValISO &"</td>"_
    &"<td style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom class=lille colspan=3><b>Med-ansv.2</b><br>Værdi "& basisValISO &"</td>"_
    &"<td style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom class=lille colspan=3><b>Med-ansv.3</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px;"" valign=bottom><b>Kode</b></td>"_
    &"</tr>"
	
    end if
    
	select case viskun
	case 1
	strViskunKri = " AND jobstatus = 1 "
	case 2
	strViskunKri = " AND (jobstatus = 1 OR jobstatus = 3)" 'alle aktive + tilbud
	case 3
	strViskunKri = " AND jobstatus = 3 " 'tilbud
	end select
	

	kvotient = 10000
	
    jobnrSQLkri = replace(jobnrSQLkri, "tjobnr", "jobnr")
    
  

	strSQL = "SELECT job.id AS jobid, jobnr, jobnavn, jobstartdato, jobslutdato, jobTpris, fakturerbart, budgettimer, fastpris, jobstatus, sandsynlighed, jo_udgifter_ulev, "_
    &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_
    &" m1.mnavn AS m1navn, m1.mnr AS m1mnr, m1.init AS m1init, "_
    &" m2.mnavn AS m2navn, m2.mnr AS m2mnr, m2.init AS m2init, "_
    &" m3.mnavn AS m3navn, m3.mnr AS m3mnr, m3.init AS m3init, "_
    &" m4.mnavn AS m4navn, m4.mnr AS m4mnr, m4.init AS m4init, "_
    &" m5.mnavn AS m5navn, m5.mnr AS m5mnr, m5.init AS m5init, "_
    &" kkundenavn, kkundenr "_
    &" FROM job"_
    &" LEFT JOIN kunder k ON (k.kid = jobknr) "_
    &" LEFT JOIN medarbejdere m1 ON (m1.mid = jobans1) "_
    &" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2) "_
    &" LEFT JOIN medarbejdere m3 ON (m3.mid = jobans3) "_
    &" LEFT JOIN medarbejdere m4 ON (m4.mid = jobans4) "_
    &" LEFT JOIN medarbejdere m5 ON (m5.mid = jobans5) "_
    &" WHERE ("& jobnrSQLkri &")  AND jobstartdato BETWEEN '"&strYear&"/"& strMrd &"/1"&"' AND '"&strYearEnd&"/"& strMrdEnd &"/"& strDayEnd &""&"' GROUP BY job.id ORDER BY jobstartdato, jobnavn" 

    'Response.Write strSQL
    'Response.flush

    'strExport = "xx99123sy#z"
    lastMonth = 0
	mdTot = 0		
			
			
    
    oRec.open strSQL, oConn, 3
	
	y = 1
	while not oRec.EOF
	select case right(y, 1)
	case 2,4,6,8,0
	bgcolthis = "#EFF3ff"
	case else
	bgcolthis = "#FFFFFF"
	end select
	
    if media <> "export" then 

	select case y
	case 1, 47, 93, 126, 172
	col = "#0083D7"
	case 2, 48, 94, 127, 173
	col = "#0099FF"
	case 3, 49, 95, 128, 714
	col = "Dimgray"
	case 4, 50, 96, 129, 175
	col = "#003399"
	case 5, 51, 97, 130, 176
	col = "#006600"
	case 6, 52, 98, 131, 177
	col = "#006699"
	case 7, 53, 99, 132, 178
	col = "#34A4DE"
	case 8, 54, 100, 133, 179
	col = "#B4E2FF"
	case 9, 55, 101, 134, 180
	col = "#DEFFFF"
	case 10, 56, 102, 135, 181
	col = "#FFCCFF"
	case 11, 57, 103, 136,182
	col = "#CCCCFF"
	case 12, 58, 104, 137, 183
	col = "#996600"
	case 13, 59, 105, 138,184
	col = "#CC9900"
	case 14, 60, 106, 139,185
	col = "#FFFF99"
	case 15, 61, 107, 140,186
	col = "#333366"
	case 16, 62, 108, 141,187
	col = "#FFFF00"
	case 17, 63, 109, 142,188
	col = "#FFDB9D"
	case 18, 64, 110, 143,189
	col = "#FFCC66"
	case 19, 65, 111, 144,190
	col = "#FF9933"
	case 20, 66, 112, 145,191
	col = "#FF794B"
	case 21, 67, 113, 146,192
	col = "#FF3300"
	case 22, 68, 114, 147,193
	col = "#990000"
	case 23, 69, 115, 148,194
	col = "#9999FF"
	case 24, 70, 116, 149,195
	col = "#6666CC"
	case 25, 71, 117, 150,196
	col = "#9999CC"
	case 26, 72, 118, 151,197
	col = "#666699"
	case 27, 73, 119, 152,198
	col = "#FFCC00"
	case 28, 74, 120, 153,199
	col = "#009900"
	case 29, 75, 121, 154,200
	col = "#66CC33"
	case 30, 76, 122, 155,201
	col = "Silver"
	case 31, 77, 123, 156,202
	col = "LightGrey"
	case 32, 78, 125, 157,203
	col = "DarkGray"
	case 33, 79, 158,204
	col = "Olive"
	case 34, 80, 159,205
    col = "#0066CC"
	case 35, 81, 160,206
	col = "#99FF66"
	case 36, 82, 161,207
	col = "#CCFFCC"
	case 37, 83, 162,208
	col = "RosyBrown"
	case 38, 84, 163,209
	col = "Honeydew"
	case 39, 85, 164,210
	col = "CornflowerBlue"
	case 40, 86, 165,211
	col = "DarkMagenta"
	case 41, 78, 166,212
	col = "skyblue"
	case 42, 88, 167,213
	col = "Thistle"
	case 43, 89, 168,214
	col = "Tomato"
	case 44, 90, 169,215
	col = "Gold"
	case 45, 91, 170,216
	col = "yellow"
	case 46, 92, 171,217
	col = "black"
	case else

        select case right(y,1)
        case 1
	    col = "#5582d2"
        case 2
	    col = "Olive"
	    case 3
        col = "#0066CC"
	    case 4
	    col = "#99FF66"
	    case 5
	    col = "#CCFFCC"
	    case 6
	    col = "RosyBrown"
	    case 7
	    col = "Honeydew"
	    case 8
	    col = "CornflowerBlue"
	    case 9
	    col = "DarkMagenta"
	    case 0
	    col = "skyblue"
        end select
	   
	end select

    end if
    

	stmonth = month(oRec("jobstartdato"))
	endmonth = month(oRec("jobslutdato"))
	thisdays = datediff("m", oRec("jobstartdato"), oRec("jobslutdato"))
	thisbudget = pipelinevalue 'oRec("jobtpris")
	


   
        if datepart("m", oRec("jobstartdato"),2,2) <> lastMonth then

            if y <> 1 then
            strJob = strJob & "<tr bgcolor='#FFFFFF'><td colspan=5 align=right style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille><b>"& formatnumber(mdTot, 2) &"</b></td><td colspan=16 style='padding:3px; border-bottom:1px #cccccc solid;'>&nbsp;</td></tr>"
            mdTot = 0
            end if


        strJob = strJob & "<tr bgcolor='#D6dff5'><td colspan=21 style='padding:3px; border-bottom:1px #cccccc solid;'><b>"& monthname(datepart("m",oRec("jobstartdato"),2,2)) &" "& datepart("yyyy",oRec("jobstartdato"), 2,2) &"</b></td></tr>"
    
        end if

    
    
    '*** Tilbud Pipelineværdi = bruttooms - udgifter * sand *** / Ellers = bruttooms - udgifter
    if oRec("jobstatus") = 3 then
    pipelinevalue = formatnumber((oRec("jobtpris") - oRec("jo_udgifter_ulev")) * (oRec("sandsynlighed")/100),2)
    else
    pipelinevalue = formatnumber((oRec("jobtpris") - oRec("jo_udgifter_ulev")),2)
    end if 


	if cint(datepart("yyyy",oRec("jobstartdato"))) <= cint(strYear) AND cint(datepart("yyyy",oRec("jobslutdato"))) >= cint(strYear) then
	strJob = strJob & "<tr bgcolor="& bgcolthis &"><td valign=top style='padding:2px; border-right:1px #cccccc solid;' class=lille><input type='hidden' name='FM_colthis_"&y&"' value='"&col&"'>"
	
	if media <> "print" AND level <=2 OR level = 6 then
	strJob = strJob & "<a href='jobs.asp?menu=job&func=red&int=1&id="&oRec("jobid")&"&rdir=pipe&nomenu=1' class='rmenu' target=""_blank"">" & left(oRec("jobnavn"), 35) & " ("& oRec("jobnr")  &")</a>"
	else
	strJob = strJob & "<b>"& left(oRec("jobnavn"), 35)  & "</b> ("& oRec("jobnr")  &")"
	end if

    strExport = strExport & oRec("jobnavn") & ";"& oRec("jobnr") &";"
	
    Select case oRec("jobstatus")
    case 3 
    strJob = strJob & " (Tilbud)"
    sandsynlighedTxt = oRec("sandsynlighed") & "%"
    case 0
    strJob = strJob & " (Lukket)"
    sandsynlighedTxt = ""
    case 2
    strJob = strJob & " (Passivt)"
    sandsynlighedTxt = ""
    case else
    strJob = strJob & ""
    sandsynlighedTxt = ""
    end select
    
    strJob = strJob & "<br>" & left(oRec("kkundenavn"), 15) & " ("& oRec("kkundenr") &")"

	strJob = strJob & "</td>" 

   strExport = strExport & oRec("kkundenavn") & ";"& oRec("kkundenr") &";"
	
	strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & formatnumber(oRec("jobtpris"), 2) &"</td>"_
	&"<td valign=top align=right style='padding:2px; white-space:nowrap;' class=lille>"& formatdatetime(oRec("jobstartdato"), 2) &" ("& thisdays &") </td>"_
    &"<td valign=top align=right style='padding:2px;' class=lille>"& sandsynlighedTxt &"</td>"_
	&"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>"& pipelinevalue &"</td>"
    
    
    strExport = strExport & formatnumber(oRec("jobtpris"), 2) &";"& formatdatetime(oRec("jobstartdato"), 2) &";"& sandsynlighedTxt &";"& pipelinevalue & ";" 
    
     
    if oRec("jobans1") <> 0 then


    if len(trim(oRec("m1init"))) <> 0 then
    uNit1 = oRec("m1init")
    else
    uNit1 = "("& oRec("m1mnr") &")"
    end if

    strJob = strJob &" <td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& uNit1 &"</td>"

    jobans1proc = oRec("jobans_proc_1")
    if jobans1proc <> 0 then
    jobans1val = pipelinevalue * (jobans1proc/100)
    else
    jobans1val = 0
    end if

  
    strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("jobans_proc_1") &"%</td>"_
    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(jobans1val, 2) &"</td>"

    strExport = strExport & uNit1 &";"& jobans1proc &";"& jobans1val & ";"

    else
        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
    
    strExport = strExport & ";;;"
    end if

    
    
    if oRec("jobans2") <> 0 then

    if len(trim(oRec("m2init"))) <> 0 then
    uNit2 = oRec("m2init")
    else
    uNit2 = "("&oRec("m2mnr") &")"
    end if


    strJob = strJob &"<td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& uNit2 &"</td>"

    
    jobans2proc = oRec("jobans_proc_2")
    if jobans2proc <> 0 then
    jobans2val = pipelinevalue * (jobans2proc/100)
    else
    jobans2val = 0
    end if

   
    strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("jobans_proc_2") &"%</td>"_
    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(jobans2val, 2) &"</td>"

    strExport = strExport & uNit2 &";"& jobans2proc &";"& jobans2val & ";"

    else
        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
    
    strExport = strExport & ";;;"
    end if


    if oRec("jobans3") <> 0 then

    if len(trim(oRec("m3init"))) <> 0 then
    uNit3 = oRec("m3init")
    else
    uNit3 = "("&oRec("m3mnr") &")"
    end if

    strJob = strJob &"<td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& uNit3 &"</td>"

    
    jobans3proc = oRec("jobans_proc_3")
    if jobans3proc <> 0 then
    jobans3val = pipelinevalue * (jobans3proc/100)
    else
    jobans3val = 0
    end if

   
    strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("jobans_proc_3") &"%</td>"_
    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(jobans3val, 2) &"</td>"

    strExport = strExport & uNit3 &";"& jobans3proc &";"& jobans3val & ";"

    else
        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
        strExport = strExport & ";;;"
    end if

    
    
    if oRec("jobans4") <> 0 then 
    
    
    if len(trim(oRec("m4init"))) <> 0 then
    uNit4 = oRec("m4init")
    else
    uNit4 = "("&oRec("m4mnr") &")"
    end if

    strJob = strJob &"<td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& uNit4 &"</td>"


    jobans4proc = oRec("jobans_proc_4")
    if jobans4proc <> 0 then
    jobans4val = pipelinevalue * (jobans4proc/100)
    else
    jobans4val = 0
    end if
    
    strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("jobans_proc_4") &"%</td>"_
    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(jobans4val, 2) &"</td>"
    
    strExport = strExport & uNit4 &";"& jobans4proc &";"& jobans4val & ";"

    else
        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
        strExport = strExport & ";;;"
    end if


    if oRec("jobans5") <> 0 then 

    if len(trim(oRec("m5init"))) <> 0 then
    uNit5 = oRec("m5init")
    else
    uNit5 = "("&oRec("m5mnr") &")"
    end if

    strJob = strJob &"<td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& uNit5 &"</td>"

    jobans5proc = oRec("jobans_proc_5")
    if jobans5proc <> 0 then
    jobans5val = pipelinevalue * (jobans5proc/100)
    else
    jobans5val = 0
    end if


    strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("jobans_proc_5") &"%</td>"_
    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(jobans5val, 2) &"</td>"

    strExport = strExport & uNit5 &";"& jobans5proc &";"& jobans5val & ";"

    else
        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
        strExport = strExport & ";;;"
    end if

    strJob = strJob &"<td bgcolor="& col &" style='padding:2px 10px 2px 2px;'>&nbsp;</td></tr>"

    if media <> "print" then
    strJob = strJob  &"<tr><td colspan=21 bgcolor=""#cccccc"" style=""height:1px;""><img src=""../ill/blank.gif"" width=""30"" height=""1"" alt="""" border=""0"" style=""border:0px;""></td></tr>"
    end if

    strExport = strExport & "xx99123sy#z"
	
    lastMonth = datepart("m", oRec("jobstartdato"),2,2) 


	
	thispipeline = formatnumber(pipelinevalue/kvotient,0) '(thisbudget/(thisdays+1))

    thisheight = thispipeline 'formatnumber(((thisbudget/(thisdays+1))/kvotient),0)
	if thisheight <> 0 then
	thisheight = thisheight
	else
	thisheight = 0
	end if
	%>
	
	<!-- januar -->
	<%

    for m = 0 to 11

    

	if (cdate(oRec("jobstartdato")) <= cdate(""& strDayno(m) &"/"& strMrdno(m)  &"/"& strYearno(m))) AND (cdate(oRec("jobstartdato")) >= cdate("1/"& strMrdno(m) &"/"&strYearno(m))) then
	tableHTML(m) = tableHTML(m) & "<tr><td><div id='j"&y&"' name='j"&y&"' style='position:relative;visibility:visible; background-color:"& col &"; border:1px "& col &" dashed;  height="&thisheight&" z-index:1000;'><img src='ill/blank.gif' width='30' height="&thisheight&" alt='"& oRec("jobnr") &"&nbsp;" & oRec("jobnavn") &"  " &  formatcurrency(thispipeline,2) &"' border='0'></div></td></tr>"
	mdIdTot(m) = mdIdTot(m) + cdbl(pipelinevalue)
	else
	Response.write "<div id='j"&y&"' name='j"&y&"' style='position:relative;visibility:hidden;  height:0; width:0; z-index:1000; display:none;'></div>"
	end if
    
    next
    %>
	

	
	
	
	<%mdTot = mdTot + pipelinevalue
    
	y = y + 1
	end if
	
	oRec.movenext
	wend
	oRec.close
	
    if y <> 1 then
    strJob = strJob & "<tr bgcolor='#FFFFFF'><td colspan=5 align=right style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille><b>"& formatnumber(mdTot, 2) &"</b></td><td colspan=16 style='padding:3px; border-bottom:1px #cccccc solid;'>&nbsp;</td></tr>"
    end if


	
	'strBudget = strBudget & "</table>" 
	strJob = strJob & "</table>" 
	
    for m = 0 to 11

    tableHTML(m) = tableHTML(m) & "</table>"

    next

     %>
	
	<%
    if media <> "export" then 
    
    '*** Pipeline table ****%>

    <%for m = 0 to 11 
    
     select case right(m, 1)
     case 0,2,4
     bgt = "whitesmoke"
     case else
     bgt = "#FFFFFF"
     end select

    %>
	<td bgcolor=<%=bgt%> valign="bottom"><%=tableHTML(m)%></td>
	<%next %>
	
	
	<%
	if media <> "print" then
	clthis = "mlille"
	else
	clthis = "lille-print"
	end if
	%>
	
    </tr>
	<tr>
	 <%for m = 0 to 11 
    
     select case right(m, 1)
     case 0,2,4
     bgt = "whitesmoke"
     case else
     bgt = "#FFFFFF"
     end select
     
     if mdIdTot(m) <> 0 then
     mdIdTot(m) = cdbl(mdIdTot(m))/kvotient
     else
     mdIdTot(m) = 0
     end if
     
     %>
	<td bgcolor=<%=bgt%> valign="bottom" align=right class=lille><%=formatnumber(mdIdTot(m), 0) %> x10k</td>
	<%next %>
	
	</tr>
	</table><br><br />
	<font class=megetlillesort>Kvo. <%=kvotient%>, antal job og tilbud: <%=y-2%></font>
	
    
    </td><td valign=top style="width:794px;">
  
	
	<% '*** job table ***
	if media <> "print" then
	bgcoltable = "#d6dff5"
	else
	bgcoltable = "#ffffff"
	end if
    %>
    <table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="<%=bgcoltable%>">
	<tr><td valign="top"><%=strJob%></td>
	</tr>
	</table>

    </td>
    </tr>
    </table>
	
     </div><!-- table div -->
	
	
	<br><br>
	<br>
	<%
    
    
    if media <> "print" then%>
	<br /><a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br><br>
    <!--
	<div style="position:relative; width:420px; background-color:#eff3ff; top:0px; left:0px; visibility:visible; border:1px red dashed; border-right:2px red dashed; border-bottom:2px red dashed; padding:15px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
		Klik på jobnavnet til venstre for at for vist hvilke blokke er hører til dette job.<br>
		Beløb bliver beregnet udfra det angivede budget / antal måneder jobbet løber over.<br>
		Hold musen over de markedrede blokke for at få vis den månedlige forventede omsætning på dette job.

        Jobstartdato,
        Jobansvarlige 1-5
        Jobstatus
        Pipelineværdi



		</div>
        -->
		<br><br>&nbsp;
	<%end if

    end if 'export



    '******************* Eksport **************************' 

if media = "export" then

    
  call TimeOutVersion() 

  
    ekspTxt = strExport
	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
    datointerval = request("datointerval")
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\pipeline.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				
				strOskrifter = "Job;Job Nr;Kontakt;Knr;Brutto Oms "& basisValISO &";Startdato;Sandsynlighed %;Pipelineværdi "& basisValISO &";Jobansv.;Andel %;Værdi "& basisValISO &";"_
                & "Jobejer.;Andel %;Værdi "& basisValISO &";Jobmedansv1.;Andel %;Værdi "& basisValISO &";Jobmedansv2.;Andel %;Værdi "& basisValISO &";Jobmedansv3.;Andel %;Værdi "& basisValISO &";"
				
				
				
				
				
				objF.writeLine("Periode afgrænsning: "& monthname(strMrdno(0)) &" "& strYearno(0) &" + 12 md "& vbcrlf)
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
				



end if






    if media <> "print" then

ptop = 0
pleft = 840
pwdt = 140

call eksportogprint(ptop,pleft,pwdt)

lnk = "&FM_kunde="&kundeid&"&FM_job="&jobid&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&FM_progrp="&progrp&"&seomsfor="&seomsfor&"&FM_start_mrd="&strMrd&""_
&"&FM_jobsog="&jobSogVal&"&viskunabnejob0="&viskunabnejob0&"&viskunabnejob1="&viskunabnejob1&"&viskunabnejob2="&viskunabnejob2&""_
&"&FM_kundejobans_ell_alle="&visKundejobans&"&FM_kundeans="&kundeans&"&FM_jobans="&jobans&"&FM_jobans2="&jobans2&"&FM_jobans3="&jobans3&"&FM_segment="&segment
%>

        
      <tr>
        <td align=center>
        <a href="#" onclick="Javascript:window.open('pipeline.asp?media=export<%=lnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
        </td><td><a href="#" onclick="Javascript:window.open('pipeline.asp?media=export<%=lnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport</a></td>
       </tr>
    <tr>
   <td align=center><a href="pipeline.asp?media=print<%=lnk%>" target="_blank"  class='rmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
    </td><td><a href="pipeline.asp?media=print<%=lnk%>" target="_blank" class="rmenu">Print version</a></td>
   </tr>
   
    <tr><td colspan="2"><br /><br />
        <%
            
                nWdt = 120
                nTxt = "Opret nyt job"
                nLnk = "jobs.asp?menu=job&func=opret&id=0&int=1&rdir=pipe"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>    
            
        

    </td></tr>
	
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
	
	
	</div>
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
