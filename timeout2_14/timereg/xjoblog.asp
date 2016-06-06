<%Response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/joblog_inc.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	
	
	 %>
	
	
	
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	 function checkAll(val) {


	        //FM_akttype_id_1
	        //alert(val)
	        antal_v = document.getElementById("antal_v_" + val).value
	        //alert(antal_v)

	        if (document.getElementById("chkalle_" + val).checked == true) {
	            tval = true
	        } else {
	            tval = false
	        }

	        //alert(tval)

	        for (i = 0; i < antal_v; i++)
	        //alert(i)
	            document.getElementById("FM_akttype_id_" + val + "_" + i).checked = tval
	    }
	
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	
	function clearJobsog(){
	document.getElementById("FM_jobsog").value = ""
	}
	
	
	
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	
	
	
	function showjoblogsubdiv(){
	document.getElementById("joblogsub").style.visibility = "visible"
	document.getElementById("joblogsub").style.display = "" 
	}
	
	function hidejoblogsubdiv(){
	document.getElementById("joblogsub").style.visibility = "hidden"
	document.getElementById("joblogsub").style.display = "none" 
	}
	
	function closeprogrpdiv(){
	document.getElementById("progrpdiv").style.visibility = "hidden"
	document.getElementById("progrpdiv").style.display = "none" 
	}
	
	function setshowprogrpval(){
	document.getElementById("showprogrpdiv").value = 1
	}
	
	function curserType(tis){
	document.getElementById(tis).style.cursor='hand'
	}
	
	
	function checkAllGK(val){
	
	antal_v = document.getElementById("antal_v").value 
	
	if (val == 1) {
	document.getElementById("chkalle1").checked = true
	document.getElementById("chkalle2").checked = false
	tval = true
	} else {
	tval = false
	document.getElementById("chkalle1").checked = false
	document.getElementById("chkalle2").checked = true
	}
	
    for (i = 0; i < antal_v ; i++)
        document.getElementById("FM_godkendt_"+i).checked = tval
}

function clearK_Jobans() {
    //alert("her")
    document.getElementById("cFM_jobans").checked = false
    document.getElementById("cFM_jobans2").checked = false
    document.getElementById("cFM_kundeans").checked = false
}
	
	
	</script>
	
	
	
	<% 
	
	
    '**************************************************'
	'***** Faste Filter kriterier *********************'
	'**************************************************'
		
	
		
	'*** Job og Kundeans ***
	call kundeogjobans()
	
	'*** Medarbejdere / projektgrupper '
    selmedarb = session("mid")
	
	'call medarbogprogrp("oms")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAns2SQLkri = ""
	jobAnsSQLkri = ""
	ugeAflsMidKri = ""
	'fakmedspecSQLkri = ""
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	if len(request("FM_medarb")) <> 0 then 'OR func = "export"
	
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
	
	
	media = request("media")
	
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
			     ugeAflsMidKri = " u.mid = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
			     ugeAflsMidKri = ugeAflsMidKri &  " OR u.mid = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			jobAnsSQLkri =  "AND ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "AND (" & jobAns2SQLkri &")"
			ugeAflsMidKri = " ("& ugeAflsMidKri &")"
			
	
	'Response.Write "ugeAflsMidKri" & ugeAflsMidKri
	'Response.end
	
	
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
	else
        aftaleid = 0
	end if
	
	
	
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
	
	
	
	
	
	
	if len(request("FM_akttype")) <> 0 AND len(request("FM_kunde")) <> 0 then 
	akttype_sel = request("FM_akttype")
	else
	    if request.Cookies("stat")("fm_akttype_sel") <> "" then
	    akttype_sel = request.Cookies("stat")("fm_akttype_sel")
	    else
	    call akttyper2009(6)
	    akttype_sel = aktiveTyper
	    end if
	end if
	
	vartyper = akttype_sel
	response.Cookies("stat")("fm_akttype_sel") = akttype_sel
	response.cookies("stat").expires = date + 40
	
	
	
	'Response.Write typerSQL
	
	
	'**** Alle SQL kri starter med NUL records ****'
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = -1 "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	    '**** Skjul timepriser ****'
	     if level <= 2 OR level = 6 then
	    
	        
		    if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hidetimepriser"))) <> 0 then    
		        hidetimepriser = 1
		        hidetpCHK = "CHECKED"
		        else
		        hidetimepriser = 0
		        hidetpCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hidetimepriser"))) <> 0 then
		        hidetimepriser = request.cookies("stat")("hidetimepriser")
		        if hidetimepriser = 1 then
		        hidetpCHK = "CHECKED"
		        else
		        hidetpCHK = ""
		        end if
		        
		        else
		        hidetimepriser = 0
		        hidetpCHK = ""
		        
		        end if
		        
		    end if
		
		end if
		response.cookies("stat")("hidetimepriser") = hidetimepriser
        '****'
        
    
            '**** Enheder **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hideenheder"))) <> 0 then    
		        hideenheder = 1
		        hideehCHK = "CHECKED"
		        else
		        hideenheder = 0
		        hideehCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hideenheder"))) <> 0 then
		        
		            hideenheder = request.cookies("stat")("hideenheder")
		            if hideenheder = 1 then
		            hideehCHK = "CHECKED"
		            else
		            hideehCHK = ""
		            end if
    		        
		        else
		        hideenheder = 0
		        hideehCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hideenheder") = hideenheder
		    
		    '***'
		    
		    
		    '**** Benyt Enheder **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("useenheder"))) <> 0 then    
		        useenheder = 1
		        useehCHK = "CHECKED"
		        else
		        useenheder = 0
		        useehCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("useenheder"))) <> 0 then
		        
		            useenheder = request.cookies("stat")("useenheder")
		            if useenheder = 1 then
		            useehCHK = "CHECKED"
		            else
		            useehCHK = ""
		            end if
    		        
		        else
		        useenheder = 0
		        useehCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("useenheder") = useenheder
		    '***'
		    
		    
		    '**** Vis kostpriser **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("visKost"))) <> 0 then    
		        visKost = 1
		        visKostCHK = "CHECKED"
		        else
		        visKost = 0
		        visKostCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("visKost"))) <> 0 AND level = 1 then
		        
		            visKost = request.cookies("stat")("visKost")
		            if visKost = 1 then
		            visKostCHK = "CHECKED"
		            else
		            visKostCHK = ""
		            end if
    		        
		        else
		        visKost = 0
		        visKostCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("useenheder") = useenheder
		    '***'
		    
	        
	        
	        '**** GK status og faktureret **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hidegkfakstat"))) <> 0 then    
		        hidegkfakstat = 1
		        hidegkCHK = "CHECKED"
		        else
		        hidegkfakstat = 0
		        hidegkCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hidegkfakstat"))) <> 0 then
		        
		            hidegkfakstat = request.cookies("stat")("hidegkfakstat")
		            if hidegkfakstat = 1 then
		            hidegkCHK = "CHECKED"
		            else
		            hidegkCHK = ""
		            end if
    		        
		        else
		        hidegkfakstat = 0
		        hidegkCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hidegkfakstat") = hidegkfakstat
	
	        '***'  
	        
	        
	        
   
	menu = request("menu")
	
	thisfile = "joblog"
	
	func = request("func")
	print = request("print")
	
	rdir = request("rdir")
	
	if print = "j" then
	    
	    if lto = "optionone" then
	    globalWdt = 600
	    else
	    globalWdt = 700
	    end if
	
    else
	
	    if rdir = "treg" OR rdir = "nwin" then
	    globalWdt = 915
	    else
	    globalWdt = 1160
	    end if
	
	end if
	
	if request("bruguge") = "1" then
	bruguge = 1
	brugugeCHK = "CHECKED"
	    
	    bruguge_week = request("bruguge_week") 
	    bruguge_year = request("bruguge_year") 
	    
	    'Response.Write "bruguge_week: " &  bruguge_week & "<br>"
	    'Response.Write "bruguge_year: " & bruguge_year
	    
	    stDato = "1/1/"&bruguge_year
	    
	    
	    datoFound = 0
	    
	    
	    
	    for u = 1 to 53 AND datoFound = 0
	    
	    
	    
	    if u = 1 then
	    tjkDato = stDato
	    else
	    tjkDato = dateadd("d",7,tjkDato)
	    end if
	    
	    tjkDatoW = datepart("ww", tjkDato, 2,2)
	    
	    if cint(bruguge_week) = cint(tjkDatoW) then
	    
	    wDay = datepart("w", tjkDato, 2,2)
	       
	        
	        select case wDay
	        case 1
	        tjkDato = dateAdd("d", 0, tjkDato)
	        case 2
	        tjkDato = dateAdd("d", -1, tjkDato)
	        case 3
	        tjkDato = dateAdd("d", -2, tjkDato)
	        case 4
	        tjkDato = dateAdd("d", -3, tjkDato)
	        case 5
	        tjkDato = dateAdd("d", -4, tjkDato)
	        case 6
	        tjkDato = dateAdd("d", -5, tjkDato)
	        case 7
	        tjkDato = dateAdd("d", -6, tjkDato)
	        end select
	    
	    stDaguge = day(tjkDato)
	    stMduge = month(tjkDato)
	    stAaruge = year(tjkDato)
	    
	    tjkDato_slut = dateadd("d", 6, tjkDato)
	    
	    slDaguge = day(tjkDato_slut)
	    slMduge = month(tjkDato_slut)
	    slAaruge = year(tjkDato_slut)
	    
	       
	    datoFound = 1
	    
	    end if
	    
	    next
	    
	    
	else
	    
	    if rdir = "treg" AND len(trim(request("bruguge_week"))) = 0 then
	    
	    stMduge = request("FM_start_mrd")
	    
	    'Response.Write "stMduge =" & request("FM_start_mrd")
	    'Response.flush
	    
	    stDaguge = request("FM_start_dag")
        stAaruge = right(request("FM_start_aar"),2) 
        slDaguge = request("FM_slut_dag")
        slMduge = request("FM_slut_mrd")
        slAaruge = right(request("FM_slut_aar"),2)
        
        bruguge = 1
	    brugugeCHK = "CHECKED"
    	
	    bruguge_week = datepart("ww", stDaguge&"/"&stMduge&"/"&stAaruge, 2,2)
	    bruguge_year = datepart("yyyy", stDaguge&"/"&stMduge&"/"&stAaruge, 2,2) 
        
        else
	    
	    bruguge = 0
	    brugugeCHK = ""
    	
    	if len(trim(request("bruguge_week"))) <> 0 then
    	bruguge_week = request("bruguge_week") 
	    bruguge_year = request("bruguge_year") 
	    else
    	
	    bruguge_week = datepart("ww", now, 2,2)
	    bruguge_year = datepart("yyyy", now, 2,2)
	    end if 
	    
	    end if
	
	end if
	
	'Response.Write "her:" & request("bruguge")
	'***** Slut faste var **'
	
	
	
	
	
	'**** Annuler kunde/overordnet godkender timer *** '
	if func = "opdaterliste" then
	
	  
	
	  ujid = split(request("ids"), ",")
	  uGodkendt = split(trim(request("FM_godkendt")), "#, ")
       

     for u = 0 to UBOUND(ujid)
	
	
	if trim(left(uGodkendt(u), 1)) = "1" then
	uGodkendt(u) = trim(left(uGodkendt(u), 1))
	editor = session("user")
	else
	uGodkendt(u) = 0
	editor = ""
	end if
	
				intGodkendt = request("tid") 
				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt(u) &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u)
				
				'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next
		%>
		
			
	    
	    
	  
		
	<% 
	
	dontshowDD = 1%>
	<!--#include file="inc/weekselector_s.asp"--> 
	<%
	
	'*** LINK og faste VAR ****
	strLink = "joblog.asp?rdir="&rdir&"&menu="&menu&"&FM_medarb="&thisMiduse&""_
	&"&lastFakdag="&lastFakdag&"&FM_job="&jobid&"&FM_start_dag="&strDag&""_
	&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&""_
	&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""_
	&"&FM_jobsog="&jobSogValPrint&""_
	&"&FM_aftaler="&aftaleid&""_
	&"&FM_radio_projgrp_medarb="&progrp_medarb&""_
	&"&FM_progrupper="&progrp&""_
	&"&FM_kunde="&kundeid&""_
	&"&FM_kundeans="&kundeans&""_
	&"&FM_jobans="&jobans&""_
	&"&FM_kundejobans_ell_alle="&visKundejobans&""_
	&"&FM_akttype="&vartyper&""_
	&"&FM_orderby_medarb="&fordelpamedarb&""_
	&"&bruguge="&bruguge&"&bruguge_week="&bruguge_week&"&bruguge_year="&bruguge_year

	'Response.Write strLink
	'Response.end
	Response.redirect strLink
	
	
	end if
	
	
	

    if print <> "j" AND rdir <> "treg" AND rdir <> "nwin" AND media <> "export" then
	pleft = 20
	ptop = 132
	jobbgcol = "#ffffff"
	else
	pleft = 20
	ptop = 0
	jobbgcol = "#ffffff"
	end if
	
	if print <> "j" AND rdir <> "treg" AND rdir <> "nwin" AND media <> "export" then
	%>
	
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	
	<%select case level
	case 1,2,6 %>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	<%end select%>
	
	<%
	
	else
	
	level = session("rettigheder") '*** Sættes normalt i topmenu **'
	
	    if print = "j" OR media = "export" then
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	    <% 
	    else
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <% 
	    
	    end if
	end if
	
	
	
	'*** Var. til komrpimeret visning ***'
        if len(request("FM_komprimeret")) <> 0 then
            komprimer = request("FM_komprimeret")
            if komprimer = "1" then 
            komCHK1 = "CHECKED"
            komCHK0 = ""
            else
            komCHK1 = ""
            komCHK0 = "CHECKED"
            end if
            Response.Cookies("stat")("komprimerliste") = komprimer
        else
            if request.Cookies("stat")("komprimerliste") = "1" then
             komCHK1 = "CHECKED"
             komCHK0 = ""
             komprimer = 1
            else
            komprimer = 0
            komCHK1 = ""
            komCHK0 = "CHECKED"
            end if
        end if
        
	    
	    '**** Vis ugeseddel ell. joblog ***'
	    if rdir = "treg" then
	        joblog_uge = 2 '** Viser ugeseddel **'
	    else
	    
	        if len(request("joblog_uge")) <> 0 then
			joblog_uge = request("joblog_uge")
			    
			    
			else
			        if request.Cookies("stat")("joblog_uge") <> "" then
			        joblog_uge = request.Cookies("stat")("joblog_uge")
			        else
			        joblog_uge = 1
			        end if
			
			end if
		
		
			
			
			    if cint(joblog_uge) = 1 then
			    joblog_ugeCHK1 = "CHECKED"
			    joblog_ugeCHK2 = ""
			    
			    joblogsubWZB = "visible"
			    joblogsubDSP = ""
			    
			    else
			    joblog_ugeCHK1 = ""
			    joblog_ugeCHK2 = "CHECKED"
			    
			    joblogsubWZB = "hidden"
			    joblogsubDSP = "none"
			    
			    end if
			
			response.Cookies("stat")("joblog_uge") = joblog_uge
	
	    end if 
	
	
	
	'**** Vis subtotaler pr akt ell. medarb. ***'
	if cint(joblog_uge) = 1 then		
	
	if len(trim(request("FM_orderby_medarb"))) <> 0 AND request("FM_orderby_medarb") <> "0" then
	fordelpamedarb = request("FM_orderby_medarb")
	    
	    if fordelpamedarb = 1 then
	    orderByKri = "Tjobnavn, Tjobnr, Tmnavn"
	    else
	    orderByKri = "Tjobnavn, Tjobnr, fase, TAktivitetId"
	    end if
	
	else
	fordelpamedarb = 0
	orderByKri = "Tjobnavn, Tjobnr"
	end if
	
	
	orderByKri = orderByKri & ", Tdato DESC"
	
	else
	
	'** Timeseddel altid pr. medarb, tdato ***'
	orderByKri = orderByKri & " Tmnr, Tdato DESC"
	komprimer = 1
	
	end if
	
	
	
	
	
	
	
	
	 
	%>
    
    
    





<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
<!-------------------------------Sideindhold------------------------------------->


    <%if print <> "j" AND media <> "export" then%>



	<% 
    
    
    oimg = "ikon_stat_joblog_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Joblog & Ugesedler"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	if rdir = "treg" then
	%>
	<a href="timereg_akt_2006.asp?showakt=1" class="vmenu"><< Tilbage</a><br /><br />
	<%
	else
	%>
	<br /><br />
	<%
	end if
	
    	    
   
        call filterheader(0,0,904,pTxt)%>
    	    
	  
	    <table cellspacing=0 cellpadding=10 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="joblog.asp?rdir=<%=rdir %>" method="post">
	    
	<%end if %> <!-- Print -->
	    
	    <%call medkunderjob %>
	    
	    </td>
	    </tr>
	    
	    
	
	<%if print <> "j" AND media <> "export" then%>
	
	<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td valign=top style="border-bottom:1px solid #cccccc;">
			<br /><b>Vis periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			
			<br /><br />
			<b>Ell. uge</b><br /> 
			
			
            <input id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> /> 
			<select name="bruguge_week">
			<% for w = 1 to 52
			
			if w = 1 then
			useWeek = 1
			else
			useWeek = dateadd("ww", 1, useWeek)
			end if
			
			if cint(datepart("ww", useWeek, 2,2)) = cint(bruguge_week) then
			wSel = "SELECTED"
			else
			wSel = ""
			end if
			
			%>
			<option value="<%=datepart("ww", useWeek, 2,2) %>" <%=wSel%>><%=datepart("ww", useWeek, 2,2) %></option>
			<%
			next %>
			
			</select>
			
			
			<select name="bruguge_year">
			<%for y = 1 to 10
			
			if y = 1 then
			useYear = dateadd("yyyy", -5, now)
			else
			useYear = dateadd("yyyy", 1, useYear)
			end if 
			
			if cint(datepart("yyyy", useYear, 2,2)) = cint(bruguge_year) then
			ySel = "SELECTED"
			else
			ySel = ""
			end if
			
			%>
			<option value="<%=datepart("yyyy", useYear, 2,2) %>" <%=ySel %>><%=datepart("yyyy", useYear, 2,2) %></option>
			<%
			
			next %>
			
			</select><br /><br />
			
			
		<%if rdir <> "treg" then %>
		<b>Visning:</b> (præsentation af data)<br />	
        <input id="joblog_uge1" name="joblog_uge" type="radio" value="1" <%=joblog_ugeCHK1 %> onclick=showjoblogsubdiv() /> <b>Vis som joblog</b> (en lang liste opdelt på job)<br />
	    
	    <div id="joblogsub" style="position:relative; background-color:#ffffff; top:0px; left:0px; width:255px; border:1px #cccccc dashed; padding:10px; visibility:<%=joblogsubWZB %>; display:<%=joblogsubDSP %>;">
	    Vis Subtotaler:<br />
        
        <%if fordelpamedarb = 1 then
		chkFordelmedarb = "CHECKED"
		else
		chkFordelmedarb = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="FM_orderby_medarb" value="1" <%=chkFordelmedarb%>>Fordél på medarbejdere<br>
	    
	    <%if fordelpamedarb = 2 then
		chkFordelpaakt = "CHECKED"
		else
		chkFordelpaakt = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio1" value="2" <%=chkFordelpaakt%>>Fordél på aktivit. (og faser)<br>
	    
	    <%if fordelpamedarb = 0 then
		chkFordelingen = "CHECKED"
		else
		chkFordelingen= ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio2" value="0" <%=chkFordelingen%>>Ingen
	    
        <!--
        
        Vis Komprimeret visning<br />
                <input name="FM_komprimeret" id="FM_komprimeret" type="radio" value="1" <%=komCHK1 %> /> Ja 
			    <input name="FM_komprimeret" id="FM_komprimeret" type="radio" value="0" <%=komCHK0 %> /> Nej
	    
	    
	    -->
            <input id="FM_komprimeret" name="FM_komprimeret" value="1" type="hidden" />
	    
	    </div>
	    
	   
	   <input id="joblog_uge2" name="joblog_uge" type="radio" value="2" <%=joblog_ugeCHK2 %> onclick=hidejoblogsubdiv() /> <b>Vis som ugesedler for medarbejdere</b><br />(Ikke jobopdelt liste, opdelt på ugebasis pr. medarb.)<br /><br />
		
		
			
		<%
		end if '** rdir **' 
		
	    %>
		
		
			<b>Kolonne indstillinger</b><br />
			<input id="hidetimepriser" name="hidetimepriser" value="1" type="checkbox" <%=hidetpCHK %> /> Skjul Timepriser<br />
		<input id="hideenheder" name="hideenheder" value="1" type="checkbox" <%=hideehCHK %> /> Skjul Enheder <br />
		<input id="useenheder" name="useenheder" value="1" type="checkbox" <%=useehCHK %> /> Benyt enheder som beregnings grundlag
		
		<%if level = 1 then %>
		<br />
		<input id="visKost" name="visKost" value="1" type="checkbox" <%=visKostCHK %> /> Vis Kostpriser
		<%end if %>
		<br />
		<input id="hidegkfakstat" name="hidegkfakstat" value="1" type="checkbox" <%=hidegkCHK %> /> Skjul Godkendt- og Faktureret -status
		
		
		
			</td>
			<td style="width:300px; border-bottom:1px solid #cccccc;" valign=top>
			
		
       
       <br /><b>Vis følgende aktivitets typer:</b> (Vælg)<br />
			<table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr><td valign=top>
			<%
			call akttyper2009(5)
			%>
			</table>
			</td></tr>
			</table>
       
       
       
       
       </td>
       
       </tr>
       <tr>
       
       <td colspan=2 align=right valign=bottom>
       
       
	<input type="submit" value=" Kør >> ">
	</td>
	</tr>
	
	
	</form>
	</table>
	
	<!-- filterDiv -->
	</td></tr>
	</table>
	</div>
	<%else 
	dontshowDD = 1%>
	<!--#include file="inc/weekselector_s.asp"--> 
	<%
	level = session("rettigheder")
	end if 'print / media%>
	




<%

if media <> "export" then 
	ldTop = 200
	ldLft = 240
	else
	ldTop = 15
	ldLft = 10
	end if
	%>
	<div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br />
	&nbsp;&nbsp;&nbsp;&nbsp;Vent veligst. Der kan gå op til 20 sek...
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>
	
	</div>
	<%
	
	Response.Flush 




startDatoKriSQL = strAar &"/"& strMrd &"/"& strDag
slutDatoKriSQL = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	
	
	'*****************************************************************************************************
		'**** Job der skal indgå i omsætning, budget og forbrugstal *******************************
		'*****************************************************************************************************
		
		
		'*** valgte job ***
		call valgtejob()
		
		
		'**** Valgte aftaler *****
		call valgteaftaler()
				
				
				    'if lto = "essens" then
					'Response.Write "her"
					'Response.Write "kid "& cint(kundeid) &" jobid "& cint(jobid) &" jobans "& cint(jobans) _ 
				    '& " kans " & cint(kundeans) & " sogval "& len(trim(jobSogVal)) & " aftid "& cint(aftaleid) 
				    'end if
				
				'*** For at spare (trimme) på SQL hvis der vælges alle job alle kunder og vis kun for jobanssvarlige ikke er slået til ****
				'*** Og der ikke er søgt på jobnavn ***
				if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 _
				 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 then 
						
					
					'jobidFakSQLkri = " OR jobid <> 0 "
					jobnrSQLkri = " OR tjobnr <> 0 "
					jidSQLkri =  " OR id <> 0 "
					seridFakSQLkri = " OR aftaleid <> 0 "
						
				end if
	
	
		'**************** Trimmer SQL states ************************
		'len_jobidFakSQLkri = len(jobidFakSQLkri)
		'right_jobidFakSQLkri = right(jobidFakSQLkri, len_jobidFakSQLkri - 3)
		'jobidFakSQLkri =  right_jobidFakSQLkri
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		
		len_seridFakSQLkri = len(seridFakSQLkri)
		right_seridFakSQLkri = right(seridFakSQLkri, len_seridFakSQLkri - 3)
		seridFakSQLkri =  right_seridFakSQLkri
		
	
		

	
	if request("print") = "j" then
	%>
	Periode <%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1)%> - <%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%> </font><br>
	<%else
	
	    if media <> "export" then%>
	    <!--include file="inc/stat_submenu.asp"-->
	    <%end if%>
	<%end if%>



        <%if cint(hidegkfakstat) <> 1 then %>
       
          
           <form action="joblog.asp?func=opdaterliste&menu=<%=menu %>&rdir=<%=rdir %>" method="post">
	
	    <input type="hidden" name="FM_aftaler" id="hidden9" value="<%=aftaleid%>">
	   <input type="hidden" name="FM_job" id="hidden8" value="<%=jobid%>">
	    <input type="hidden" name="FM_jobsog" id="hidden7" value="<%=jobSogVal%>">
	   <input type="hidden" name="FM_radio_projgrp_medarb" id="hidden4" value="<%=progrp_medarb%>">
		<input type="hidden" name="FM_progrupper" id="hidden5" value="<%=progrp%>">
		<input type="hidden" name="FM_medarb" id="hidden6" value="<%=thisMiduse%>">
	    <input type="hidden" name="FM_kunde" id="hidden1" value="<%=kundeid%>">
	    <input type="hidden" name="FM_kundeans" id="hidden2" value="<%=kundeans%>">
	    <input type="hidden" name="FM_jobans" id="hidden3" value="<%=jobans%>">
	   <input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
	    <input type="hidden" name="FM_akttype" value="<%=vartyper%>">
	    <input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
	    <input type="hidden" name="bruguge" value="<%=bruguge%>">
	    <input type="hidden" name="bruguge_week" value="<%=bruguge_week%>">
	    <input type="hidden" name="bruguge_year" value="<%=bruguge_year%>">
	    
	    <input type="hidden" name="FM_start_mrd" value="<%=strMrd%>">
	    <input type="hidden" name="FM_start_dag" value="<%=strDag%>">
	    <input type="hidden" name="FM_start_aar" value="<%=strAar%>">
	    <input type="hidden" name="FM_slut_dag" value="<%=strDag_slut%>">
	    <input type="hidden" name="FM_slut_mrd" value="<%=strMrd_slut%>">
	    <input type="hidden" name="FM_slut_aar" value="<%=strAar_slut%>">
	    
	    
       
        
	    
	   <%end if %>
	   
	   
       
	    
	    
      
       


        <%
        
        dim vlgt_typerTotaler, vlgt_typerTotalerGrand, vlgt_typerTotalerMed, vlgt_typerTotalerAkt
        redim vlgt_typerTotaler(100), vlgt_typerTotalerGrand(100), vlgt_typerTotalerMed(100), vlgt_typerTotalerAkt(100)
        
        dim vlgt_typerTotalerEnh, vlgt_typerTotalerGrandEnh, vlgt_typerTotalerMedEnh, vlgt_typerTotalerAktEnh
        redim vlgt_typerTotalerEnh(100), vlgt_typerTotalerGrandEnh(100), vlgt_typerTotalerMedEnh(100), vlgt_typerTotalerAktEnh(100)
        
         
        visopdaterknap = 0

		'*** Smiley ***''
		call ersmileyaktiv()
	    
	   
	    '*** Afsluttede uger ***''
		dim medabAflugeId
		redim medabAflugeId(200)
		strSQL2 = "SELECT u.status, u.afsluttet, WEEK(u.uge, 1) AS ugenr, YEAR(u.uge) AS aar, u.id, u.mid FROM ugestatus u WHERE "& ugeAflsMidKri &" AND uge BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"' GROUP BY u.mid, uge"
		'Response.write strSQL2
		'Response.end
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
			medabAflugeId(oRec2("mid")) = medabAflugeId(oRec2("mid")) & "#"& oRec2("ugenr") &"_"& oRec2("aar") &"#,"
		oRec2.movenext
		wend
		oRec2.close 
		
	
	'*******************'
	'**** MaIN SQL *****'
	'*******************'
	
	'*** Vis alle timer hvor man er jobanssv. **'
	if cint(visKundejobans) = 1 then
	medarbSQlKri = "tmnr <> 0" 
	else
	medarbSQlKri = medarbSQlKri
	end if
	
	call akttyper2009(7)
	
	lastjobnavn = ""
	lastmedarb = 0
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, "_
	&" a.fakturerbar,"_
	&" Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom, Tknr,"_
	&" sttid, sltid, a.faktor, godkendtstatus, godkendtstatusaf, kkundenr, "_
	&" m.mnr, m.init, mnavn, m.mid, a.id AS aid, v.valutakode, t.valuta, k.kid, t.kostpris, "_
	&" a.fase, a.antalstk, a.aktbudgetsum, a.bgr, a.aktbudget "_
	&" FROM timer t "_
	&" LEFT JOIN aktiviteter a ON (a.id = TAktivitetId)"_
	&" LEFT JOIN kunder k ON (k.kid = Tknr)"_
	&" LEFT JOIN medarbejdere m ON (m.mid = tmnr)"_
    &" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
	&" WHERE ("& jobnrSQLkri &") AND ("& medarbSQlKri &") AND ("& aty_sql_sel &") AND "_
	&" (Tdato BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"') ORDER BY "& orderByKri
	
	'&" WHERE "& jobMedarbKri &" "& selaktidKri &""_
    'Response.write "kundeid:"& kundeid &"  "&  strSQL & "<br><br>"
	
	
	'Response.Write "visKundejobans "& visKundejobans &"<br>jobAnsSQLkri:" & jobAnsSQLkri & "<br>"& jobAns2SQLkri & "<br> kundeans"& kundeAnsSQLkri & "<br><br>"
	
	'if lto = "essens" then
	'Response.Write strSQL
	'Response.flush
	'end if
	
	oRec.open strSQL, oConn, 0, 1
	'showFaseTot = 0
	v = 0
	lastFase = ""
	thisFase = ""
	while not oRec.EOF
		
		'if isdate(oRec("Tdato")) = true then
		strWeekNum = datepart("ww", oRec("Tdato"),2,2)
		'else
		'strWeekNum = 1
		'end if
		
		
		thisFase = oRec("fase")
		id = 1
		
			
			if lto = "bowe" then
			
			
					
					'*** Akt navn (Dag / Aften / Nat) ****
					
					akttid = ""
					if instr(oRec("Anavn"), "(Nat)") <> 0 OR instr(oRec("Anavn"), "(Aften)") <> 0 OR instr(oRec("Anavn"), "(Dag)") <> 0 then 
					
					if instr(oRec("Anavn"), "(Nat)") <> 0 then
					akttid_st = instr(oRec("Anavn"), "(N")
					akttid_sl = instr(oRec("Anavn"), "t)")
					end if
					
					if instr(oRec("Anavn"), "(Aften)") <> 0 then
					akttid_st = instr(oRec("Anavn"), "(A")
					akttid_sl = instr(oRec("Anavn"), "n)")
					end if
					
					if instr(oRec("Anavn"), "(Dag)") <> 0 then
					akttid_st = instr(oRec("Anavn"), "(D")
					akttid_sl = instr(oRec("Anavn"), "g)")
					end if
					
					if akttid_st <> 0 AND akttid_sl <> 0 then
					akttid_length = (akttid_sl - akttid_st) 
					akttid_str = mid(oRec("Anavn"), akttid_st+1, akttid_length)
					akttid = akttid_str
					else
					akttid = ""
					end if
					
					
			        end if
					
					
					
					'*** KA nr ****
					kanr_st = instr(oRec("Tjobnavn"), "(")
					kanr_sl = instr(oRec("Tjobnavn"), ")")
					
					if kanr_st <> 0 AND kanr_sl <> 0 then
					kanr_length = (kanr_sl - kanr_st) 
					kanr_str = mid(oRec("Tjobnavn"), kanr_st+1, kanr_length-1)
					kanr = kanr_str
					else
					kanr = ""
					end if
					
					'** bowekode **
					bowekode_st = instr(oRec("Tjobnavn"), ")")
					
					if bowekode_st <> 0 then
					bowekode_length = 2
					bowekode = mid(oRec("Tjobnavn"), bowekode_st+1, 3)
					else
					bowekode = ""
					end if
					
					ekspTxt = ekspTxt & oRec("Tknavn")&";"&oRec("kkundenr")&";"&oRec("Tjobnavn") &";"&kanr&";"&bowekode&";"& oRec("Tjobnr")&";"
			
			else
		
			ekspTxt = ekspTxt & oRec("Tknavn")&";"&oRec("kkundenr")&";"&oRec("Tjobnavn") &";"& oRec("Tjobnr")&";"
		
			end if 'bowe
			
			
			if media <> "export" then
			
			
			'** Fordel på medarb **'
			if cint(joblog_uge) = 1 then
			
			
		
		        
		        'Response.write lastmedarb &" <> " & oRec("Tmnr") & " job: " & lastjobnavn &" <> " & oRec("Tjobnavn") 
		        'Response.Write lastaktid & "<>"& oRec("Taktivitetid")
				
		        if fordelpamedarb = 1 then
				
				    if (lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr")) then
					    if v <> 0 then
    					
					    call fordelpamedarbejdere()
    					
    					
					    end if
				    end if
				
				end if
				
				
				'** Fordel på akt **'
				if fordelpamedarb = 2 then
				
				    if (lastaktid <> oRec("Taktivitetid") OR lastjobnr <> oRec("Tjobnr")) then
					    if v <> 0 then
					    
					    '** Nulstiller fase ved næste job ***'
					    if lastjobnr <> oRec("Tjobnr") then
    					call fordelpaakt(1)
    					else
    					call fordelpaakt(0)
    					end if
    					
					    
    					
    					
					   
    					
					    end if
				    end if
				
				end if
		        
		        
				
				if lastjobnr <> oRec("Tjobnr") then
				
				
				        if v <> 0 then
        				
				        call jobtotaler()
        				
				      
				        end if
				
				
				
				 
				   
				
				tTop = 40
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)

                
                       call jobansoglastfak()
                
                
                
                %>
                <table border="0" width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
				
				
				
				<%
				komprimer = "1"
				if komprimer <> "1" then %>
				<tr>
				    
					<td colspan="16" style="padding:5px; height:100px;">
					
					
					
					
					<table cellpadding=0 cellspacing=0 border=0 width=100%>
					<tr><td valign=top>
					<img src="../ill/ikon_kunder_16.png" alt="Kontakt" border="0">&nbsp;&nbsp;<%=oRec("Tknavn")%>&nbsp;(<%=oRec("kkundenr")%>)<br>
					<img src="../ill/ikon_job_16.png" alt="Job" border="0">&nbsp;&nbsp;<b><%=oRec("Tjobnavn")%> (<%=oRec("Tjobnr")%>)</b><br />
					
					
						Jobansvarlig 1: <b><%=jobans1txt%></b>
						<%if len(trim(jobans2txt)) <> 0 then%>
						<br>
	   					Jobansvarlig 2: <b><%=jobans2txt%></b>
						<%end if
						jobans1txt = ""
						jobans2txt = ""
						%>
						</td><td valign=top style="padding-left:20px;">
						
						
						<%select case jobstatus
						case 1
						thisjobstatus = "Aktiv"
						case 2
						thisjobstatus = "Passiv"
						case 0
						thisjobstatus = "Lukket"
						end select%>
						Jobstatus: <b><%=thisjobstatus%></b><br>
						Jobperiode: <b><%=formatdatetime(jobstartdato, 1)%> - <%=formatdatetime(jobslutdato, 1)%></b>
						<br>
						Forkalkuleret timer: <b><%=formatnumber(jobForkalkulerettimer, 2)%></b><br>
						
						<%if fastpris = 1 then %>
						Jobtype: <b>Fastpris</b>
						
						<%if cint(usejoborakt_tp) <> 0 then   %>
						(aktiviteter)
						<% else %>
						(job)
						<%end if %>
						
						<%else %>
						Jobtype: <b>Lbn. Timer</b>
						<%end if %><br />
						
						<%if len(trim(rekvnr)) <> 0 then %>
						Rekvisitions nr: <b><%=rekvnr %></b><br />
						<%end if %>
						
						Sidste fakturadato: 
						<%if formatdatetime(lastfakdato, 2) <> "01-01-2001" then%>
						<b><%=formatdatetime(lastfakdato, 2)%></b>
						<%else%>
						-
						<%end if%>
						
						</td></tr>
						</table>
					    
						
					</td>
					
				</tr>
				<%end if '*** End Komprimer **'
				
				call tableheader()
				
				
				
				end if '*** job <> lastjob **'
				
				
				
		else '** joblog_uge = 2 **'
		
		            
		            
		            'Response.Write "lastdate: " & cdate(lastdate) &" <> "& cdate(oRec("tdato")) & " v: "& v &"<br>" 
		            
		            if cdate(lastdate) <> cdate(oRec("tdato")) OR (lastmedarb <> oRec("Tmnr"))  then
		            
		             if v <> 0 then
		                    
		                call dagstotaler(2)
		                
		                timerTotdag = 0
				        enhederTotdag = 0
				     else
				        
				        call dagstotaler(1)
				        
		             end if
		             
		           
		            
		            end if
		            
				
		            
		            
		            
		            if lastWeekNum <> strWeekNum OR lastmedarb <> oRec("Tmnr") then
				    
				    
				    
				   if v <> 0 then
				   
				
        				
				        call jobtotaler()
        				
				       
				   
				   end if
				    
				
				tTop = 40
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)



                %>
                
                <!--lastWeekNum &" <> "& strWeekNum &" OR "& lastmedarb &" <>"& oRec("Tmnr")  -->
                
                
                <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
              
				    
					
				    <tr>
				        <td colspan=16 bgcolor="#ffffff" style="padding:10px 10px 0px 10px;">
				        <h4>Uge: <%=datepart("ww", oRec("tdato"), 2,2)&" - "& datepart("yyyy", oRec("tdato"), 2,2)%> - <%=oRec("tmnavn") %> (<%=oRec("tmnr") %>)</h4>
				        
				        <%
				        call erugeAfslutte(datepart("yyyy", oRec("tdato"), 2,2), datepart("ww", oRec("tdato"), 2,2), oRec("tmnr")) 
				        %>
				        
				        <%if showAfsuge = 0 then %>
				        Denne uge er afsluttet via smiley d. <%=formatdatetime(cdAfs, 2)%> kl. <%=formatdatetime(cdAfs, 3)%> (uge <%=datepart("ww", cdAfs, 2, 2)%>)
                        
		                <%end if%>
				        
				        </td>
				    </tr>
				    
				    
				    
				    <%
				    
				    
				    call tableheader()
                    
                    call dagstotaler(0)
                                
				    end if
				    
				    
				
				
				
				
		end if '*** End Joblog_uge **'
		
		end if '** media	  
				
		        
		         '*** Finder jobans og last fak hvis liste vises    ***'
                 '*** som ugeseddel, ellers er                      ***'
                 '*** jobans og last fak allrrede fundet            ***'
                 if cint(joblog_uge) = 2 then
                 call jobansoglastfak()
                 end if 
                 
                 '***'
	   
	   
	         
				
				
				  if (cint(oRec("tfaktim")) = 20 OR cint(oRec("tfaktim")) = 21 OR _
				   cint(oRec("tfaktim")) = 8 OR cint(oRec("tfaktim")) = 81) AND _
				    level <> 1 AND cint(session("mid")) <> cint(oRec("tmnr")) then 
	               hidesygtyper = 1
	               else 
	               hidesygtyper = 0
	              end if
	              
	              
	              '*** Skjul sygdmstyper for andre end end selv + admin ****
	             if hidesygtyper <> 1 then 
				
				
				if media <> "export" then
				%>
				
				
				
				
				
				<tr>
					<td bgcolor="#d6DFf5" colspan="16"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<tr>
				<td valign="top">&nbsp;</td>
				<td style="padding-top:3px; width:25px;" valign="top"><%=strWeekNum%>
				</td>
				<td align=right style="padding-top:3px; width:80px; padding-right:10px;" valign="top"><%=left(weekdayname(weekday(oRec("Tdato"))), 3) %>. <%=day(oRec("Tdato")) &" "&left(monthname(month(oRec("Tdato"))), 3) &". "& right(year(oRec("Tdato")), 2)%></td>
				
				
				<%if komprimer = "1" then %>
				<td style="padding-top:3px; padding-right:15px; width:175px;" valign="top">
				<b><%=oRec("Tknavn")%></b>&nbsp;(<%=oRec("kkundenr")%>)<br />
				
				<%=oRec("tjobnavn")%> (<%=oRec("tjobnr") %>)
				
				<font class=megetlillesilver>
				
				<%if cint(fastpris) = 1 then %>
				(fastpris)
				<%end if %>
				
				
				<%if jobans1 <> 0 then %>
				<br /><%=jobans1txt %>
				<%end if %>
				
				<%if jobans2 <> 0 then %>
				        
				        <%if jobans1 <> 0 then %>
				        <%=", " %>
				        <%end if %>
				        <%=jobans2txt %>
				<%end if %>
				</font>
				
				</td>
				<%else %>
				<td valign="top" style="width:10px;">
                    &nbsp;</td>
				<%end if %>
				
				<td valign=top style="padding:3px 10px 5px 0px;"><%=oRec("fase") %>&nbsp;</td>
				
				<%
				
				end if 'media
				
				ekspTxt = ekspTxt & oRec("fase") & ";"
				 
				ekspTxt = ekspTxt & strWeekNum &";"& formatdatetime(oRec("Tdato"), 2) &";"
				
                '*** Er uge afsluttet af medarb, er smiley og autogk slået til ***'
                erugeafsluttet = instr(medabAflugeId(oRec("mid")), "#"&datepart("ww", oRec("Tdato"),2,2)&"_"&datepart("yyyy", oRec("Tdato"))&"#")
                 
                 
                
              
                 if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) = year(now) AND DatePart("m", oRec("Tdato")) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("Tdato")) > 1))) then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
				
				'if print <> "j" then
				'awdt = 350
				'else
				awdt = 150
				'end if
				
				call akttyper(oRec("fakturerbar"), 1)
				if media <> "export" then
				%>
				<td style="padding:3px 10px 5px 0px; width:<%=awdt%>px;" valign="top">
				<%
				    if ((oRec("godkendtstatus") = 0 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 0) _
				    OR (oRec("godkendtstatus") = 0 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 1 AND level = 1)) _
				    AND (cdate(lastfakdato) <  cdate(oRec("Tdato"))) then %>
    					
    					
    					
					    <a href="rediger_tastede_dage_2006.asp?id=<%=oRec("Tid")%>" target="_blank" class=vmenu>
    					
					    <%=oRec("Anavn")%>
    					
					    <img src="../ill/blyant.gif" width="12" height="11" alt="Tilføj kommentar" border="0">
					    </a>
					    
				        &nbsp;<font class=megetlillesort>(<%=akttypenavn%>)</font>
    				    
    					
    					
    					
    				
				    <%else%>
					    <b><%=oRec("Anavn")%></b>
					    <%
					    
					    %>
				        &nbsp;<font class=megetlillesort>(<%=akttypenavn%>)</font> 
    				    
    					
    					
    					
				    <%end if
				
				if len(trim(oRec("timerkom"))) <> 0 then%>
				<br /><i><%=oRec("timerkom")%></i>
				<%end if%>
				
			    <%
				'**** Kundekommentarer ****
					
					kundeKomm = ""
					strSQLin = "SELECT note FROM incidentnoter WHERE timerid = "& oRec("tid") &" ORDER BY id"
					oRec3.open strSQLin, oConn, 3 
					i = 0
					
					while not oRec3.EOF 
					%>
					
					<br><br><i><%=oRec3("note")%></i>
                   
                    <%
					kundeKomm = kundeKomm & oRec3("note") & "<br /><br />"
					
					
					
					
					oRec3.movenext
					wend
					oRec3.close  
					%>
					
				
				 &nbsp;</td>
				<%
				
                
                
                end if '** media
                
                aktnavnEksp = ""
                if len(oRec("Anavn")) <> 0 then
                aktnavnEksp = replace(oRec("Anavn"), Chr(34), "''")
                else
                aktnavnEksp = ""
                end if
                
               
				call akttyper2009Prop(oRec("tfaktim"))
				ekspTxt = ekspTxt & aktnavnEksp &";"& akttypenavn &";"& aty_fakbar &";" 
				
				if lto = "bowe" then
					ekspTxt = ekspTxt & akttid &";" 
				end if
				
				if media <> "export" then%>
				<td style="padding-top:3px; padding-left:5px;" valign="top"><%=left(oRec("Tmnavn"),16)%>&nbsp;(<%=oRec("mnr")%>)</td>
				<%
				end if
				
				ekspTxt = ekspTxt & oRec("mnavn") &";"&oRec("mnr")&";"&oRec("init")&";"
				
				
				if media <> "export" then%>
				<td align="right" style="padding-top:3px; padding-right:5px;" valign="top"><b><%=formatnumber(oRec("Timer"), 2)%></b>
				<%
				end if
				ekspTxt = ekspTxt & oRec("Timer") &";"
				
				
				if len(oRec("sttid")) <> 0 then
				
					if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
					    if media <> "export" then
					    Response.write "<font class=megetlillesort><br>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5)
					    end if
					    ekspTxt = ekspTxt & left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5) &";"
					else
					    ekspTxt = ekspTxt &";"
					end if
				
				else
				ekspTxt = ekspTxt &";"
				end if
				
				if media <> "export" then%>
				</td>
				<%end if
				
				
				enheder = 0
				    enheder = oRec("faktor") * oRec("timer")
    				
				    if enheder <> 0 then
				    enheder = cdbl(enheder)
				    else
				    enheder = 0
				    end if
				
				if cint(hideenheder) = 0 then 
				
				    
				    if media <> "export" then%>
				    <td align="right" style="padding-top:3px; padding-right:5px; white-space:nowrap;" valign="top"><%=formatnumber(enheder, 2)%></td>
				    <%end if
				    ekspTxt = ekspTxt & enheder &";"
				
				else
				    if media <> "export" then%>
				    <td>&nbsp;</td>
				    <%end if
				end if
				
				vlgt_typerTotaler(oRec("tfaktim")) = vlgt_typerTotaler(oRec("tfaktim")) + oRec("Timer")
				vlgt_typerTotalerAkt(oRec("tfaktim")) = vlgt_typerTotalerAkt(oRec("tfaktim")) + oRec("Timer")
				vlgt_typerTotalerMed(oRec("tfaktim")) = vlgt_typerTotalerMed(oRec("tfaktim")) + oRec("Timer")
				vlgt_typerTotalerGrand(oRec("tfaktim")) = vlgt_typerTotalerGrand(oRec("tfaktim")) + oRec("Timer")
				
				vlgt_typerTotalerEnh(oRec("tfaktim")) = vlgt_typerTotalerEnh(oRec("tfaktim")) + enheder
				vlgt_typerTotalerAktEnh(oRec("tfaktim")) = vlgt_typerTotalerAktEnh(oRec("tfaktim")) + enheder
				vlgt_typerTotalerMedEnh(oRec("tfaktim")) = vlgt_typerTotalerMedEnh(oRec("tfaktim")) + enheder
				vlgt_typerTotalerGrandEnh(oRec("tfaktim")) = vlgt_typerTotalerGrandEnh(oRec("tfaktim")) + enheder
				
				
				
				
				
				
				
				'*** Timepriser ***'
				if (level <=2 OR level = 6) AND cint(hidetimepriser) = 0 then
				
				if media <> "export" then%>
				<td style="padding-top:3px; padding-right:5px; white-space:nowrap;" align="right" valign="top">
				<%
				end if
				    
				    
				    
				    call akttyper2009Prop(oRec("tfaktim"))
				    '*** Alle fakturerbare ***'
				    
				    if aty_fakbar = 1 OR cint(oRec("tfaktim")) = 61 then 'stk antal
				    
				        if cint(fastpris) = 0 then
    				    tpris = formatnumber(oRec("TimePris"), 2)
    				    else
    				        if cint(usejoborakt_tp) = 0 then
    				        tpris = formatnumber(fasttimePris, 2) 'Fra job
    				        else
    				        tpris = formatnumber(oRec("aktbudget"), 2) 'Enheds pris timer eller. stk. fra aktivitet
    				        end if
    				    end if 
    				    
    				    
    				    if media <> "export" then%>
    				    <b><%=tpris %></b>
				        <%end if 
				    
				    ekspTxt = ekspTxt & tpris &";"
    				else
    				
    				if media <> "export" then%>
    				&nbsp;
    				<%end if
    				ekspTxt = ekspTxt & ";"
				    end if
				
				    if media <> "export" then%>
				    </td>
			        <td style="padding-top:3px; padding-right:5px; white-space:nowrap;" align="right" valign="top">
				    <%
				    end if
				    
				    if aty_fakbar = 1 OR cint(oRec("tfaktim")) = 61 then 'stk. antal%>
				        
				    <%if useenheder = 1 then
				    tprisTot = tpris*enheder
				    else
				    tprisTot = tpris*oRec("timer")
				    end if 
				    
				    if media <> "export" then%>
				    <%=formatnumber(tprisTot , 2)&" "&oRec("valutakode")%>
			        <%end if 
			        
			        
			        ekspTxt = ekspTxt & formatnumber(tprisTot, 2) &";"&oRec("valutakode")&";"
    			    else
    			    if media <> "export" then%>
    			    &nbsp;
    			    <%end if
    			    ekspTxt = ekspTxt & ";;"
				    end if
				    
				    if media <> "export" then%>
				    </td>
				    <%end if
				
				    else
				        if media <> "export" then %>
				        <td>&nbsp;</td>
				        <td>&nbsp;</td>
				        <%end if
				end if %>
				
				
				<%
				''*** Kostpriser ***'
				if level = 1 AND cint(visKost) = 1 then
				
				if media <> "export" then%>
				<td style="padding-top:3px; padding-right:5px; white-space:nowrap;" align="right" valign="top">
				<%end if
				    
				    call akttyper2009Prop(oRec("tfaktim"))
				    '*** Alle fakturerbare ***'
				   
				    if aty_fakbar = 1 then
				        if media <> "export" then
				        %>
    				    <%=formatnumber(oRec("kostpris"), 2)%>
				        <%end if 
				    ekspTxt = ekspTxt & formatnumber(oRec("kostpris"), 2) &";"
    				else
    				    if media <> "export" then%>
    				    &nbsp;
    				    <%end if
    				ekspTxt = ekspTxt & ";"
				    end if
				
				 if media <> "export" then
			     %>
				 </td>
			     <td style="padding-top:3px; padding-right:5px; white-space:nowrap;" align="right" valign="top">
				 <%
				 end if
				 
				 if aty_fakbar = 1 then%>
				    
				    <%if useenheder = 1 then
				    kostTot = oRec("kostpris")*enheder
				    else
				    kostTot = oRec("kostpris")*oRec("timer")
				    end if 
				    
				     if media <> "export" then%>
				     <%=formatnumber(kostTot, 2)&" DKK"%>
			         <%end if
			         ekspTxt = ekspTxt & formatnumber(kostTot, 2) &";DKK;"
    			     
    			     else
    			      if media <> "export" then%>
    			      &nbsp;
    			      <%end if
    			    ekspTxt = ekspTxt & ";;"
				  
				  end if 'aty_fakbar
				    
				     if media <> "export" then%>
				    </td>
				    <%
				    end if
				else 
				
				     if media <> "export" then%>
				     <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <%end if
				
				end if
				
				
				 if media <> "export" then
				    
				    %>
				     <!-- Taste dato -->
				    <%if print <> "j" then %>
				    <td style="padding-top:3px; padding-right:5px;" align="right" valign="top"><font class="megetlillesilver"><%=convertDate(oRec("TasteDato"))%></font></td>
				    <%else %>
				    <td>&nbsp;</td>
				    <%end if 
				
				end if%>
				
				<%if cint(hidegkfakstat) <> 1 then 
				    if media <> "export" then%>
				    <td class=lille valign=top align=center style="padding:3px 5px 3px 3px;">
                    <%end if
                    
                 if cint(oRec("godkendtstatus")) = 0 then
                 gkCHK = ""
                 erGk = 0
                 else
                 gkCHK = "CHECKED"
                 erGk = 1
                 end if
                 
                 erGkaf = ""
                    
                 if len(trim(jobans1)) <> 0 then
                 jobans1 = jobans1
                 else
                 jobans1 = 0
                 end if
                 
                 if len(trim(jobans2)) <> 0 then
                 jobans2 = jobans2
                 else
                 jobans2 = 0
                 end if
                 
                    
                    if media <> "export" then
                 
					if level = 1 OR _
					cint(session("mid")) = cint(jobans1) OR _
					cint(session("mid")) = cint(jobans2) then
					
					    '*** Godkendelse ***'
					     if print <> "j" then%>
    					
					     <input name="ids" id="ids" value="<%=oRec("tid")%>" type="hidden" />
					    <input type="checkbox" name="FM_godkendt" id="FM_godkendt_<%=v%>" value="1" <%=gkCHK %>>
                        <input name="FM_godkendt" id="FM_godkendt_hd_<%=v%>" value="#" type="hidden" />
					    <%
					    else
    					    
					         if cint(oRec("godkendtstatus")) = 1 then
				            %>
				            <b>Godkendt!</b>
				            <%
				           
				            end if
    					    
					    end if
					
					visopdaterknap = 1
					else
					    if cint(oRec("godkendtstatus")) = 1 then
				        %>
				        <b>Godkendt!</b>
				        <%
				        
				        end if
					end if
					
					end if 'Media
					
					if len(trim(oRec("godkendtstatusaf"))) <> 0 then
					    if media <> "export" then
					    %>
					    <br /><i><%=left(oRec("godkendtstatusaf"), 15)%></i>
					    <% 
					    end if
					erGkaf = oRec("godkendtstatusaf")
					end if 
					
			        if media <> "export" then%>
				    </td>
				    <td valign="top" style="padding:3px 5px 3px 3px;">
				    <%
				    end if
				
				erFaktureret = 0
				
				if cdate(lastfakdato) >= cdate(oRec("Tdato")) then 
				
				erFaktureret = 1%>
				    
				    <%
				     if media <> "export" then
				     
				        if (level = 1 OR _
					    cint(session("mid")) = cint(jobans1) OR _
					    cint(session("mid")) = cint(jobans2)) AND (print <> "j") then%>
					    <a href="erp_fakhist.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=jobid%>" class=vmenu target="_blank">Ja</a>
					    <%else %>
					    Ja
					    <%
					    end if
					
					end if %>
				
				
				
				<%end if 
				
				 if media <> "export" then%>
				</td>
				</tr>
				<%end if %>
				
				
				<%
				
				
				if lto <> "execon" then
				ekspTxt = ekspTxt & erGk &";"& erGkaf &";"& erFaktureret &";"
				end if
				
				else 
				     if media <> "export" then
				    %>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <%end if
				
				end if '** gkstat **'
				
				
				
				komm_note_Txt = ""
				if len(oRec("timerkom")) <> 0 then
				komm_note_Txt = oRec("timerkom") & kundeKomm
				end if
                
                
                if instr(komm_note_Txt, " -") <> 0 then
                komm_note_Txt = replace(komm_note_Txt, " -", "-")
                end if
                
                call htmlparseCSV(komm_note_Txt)
                
               if len(komm_note_Txt) <> 0 then 
               ekspTxt = ekspTxt & """ " & htmlparseCSVtxt &" "";"
               else
               ekspTxt = ekspTxt & ";"
               end if%>
				
				
				
				
				
				
			<%
			ekspTxt = ekspTxt & "xx99123sy#z" 
			
			v = v + 1
			lastmedarbnavn = oRec("Tmnavn")
			lastmedarb = oRec("Tmnr")
			lastjobnr = oRec("Tjobnr")
			lastjobnavn = oRec("Tjobnavn")
			lastaktid = oRec("Taktivitetid")
			lastaktnavn = oRec("Anavn")
			lastakttype = akttypenavn
			lastWeekNum = strWeekNum 
			lastdate = oRec("tdato")
			lastFase = oRec("fase")
			'lastYearNum = 
	        
	        
	        
	        Response.flush
	       
	       
	       end if 'hidesygtyper = 1
	       
	        oRec.movenext
	        wend
	        oRec.Close
	        'Set oRec = Nothing


        %>
         <script>
             document.getElementById("load").style.visibility = "hidden";
             document.getElementById("load").style.display = "none";
			    </script>

    <%
     if media <> "export" then
    
    if cint(joblog_uge) = 1 then
    
    
	        if v > 0 AND fordelpamedarb = 1 then
        	    call fordelpamedarbejdere()
	        end if
        	
	        if v > 0 AND fordelpamedarb = 2 then
                call fordelpaakt(1)
        	end if
	
	
	end if
	
	
	if v = 0 then
	%>
	<tr>
		<td colspan=11 valign="top" style="padding:20px; border:1px #5582d2 solid;"><br /><br /><b>Info:</b><br>
		Der blev ikke fundet registreringer der matcher de valgte kriterier.<br>
            &nbsp;</td>
	</tr>
	<%
	end if
	
	
	

end if 'media

if v <> 0 then

    if media <> "export" then

        if cint(joblog_uge) = 2 then
        call dagstotaler(3)
        end if
        
            call jobtotaler()
            call grandTotal
        
        %><br /><br /><br /><%
    
    end if


'******************* Eksport **************************' 

if media = "export" then


    hidetimepriser = request.cookies("stat")("hidetimepriser")
    hideenheder = request.cookies("stat")("hideenheder")
    hidegkfakstat = request.cookies("stat")("hidegkfakstat")
    

	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	datointerval = request("datointerval")
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\joblog.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				if lto = "bowe" then
				strOskrifter = "Kontakt;Kontakt Id;Job;KA Nr;Böwe kode;Job Nr;Uge;Dato;Aktivitet;Type;Fakturerbar;Akt. tidslås;Medarbejder;Medarb. Nr;Initialer;Antal;Tid/Klokkeslet;"
				else
				strOskrifter = "Kontakt;Kontakt Id;Job;Job Nr;Fase;Uge;Dato;Aktivitet;Type;Fakturerbar;Medarbejder;Medarb. Nr;Initialer;Antal;Tid/Klokkeslet;"
				end if
				
				if cint(hideenheder) = 0 then
				strOskrifter = strOskrifter &"Enheder;"
				end if
				
				
				if (level <=2 OR level = 6) AND cint(hidetimepriser) = 0 then
				strOskrifter = strOskrifter & "Timepris;Timepris ialt;Valuta;"
				end if 
				
				if level = 1 AND visKost = 1 then 
				strOskrifter = strOskrifter & "Kostpris;Kostpris ialt;Valuta;"
				end if
				
				if lto <> "execon" AND cint(hidegkfakstat) = 0 then
				strOskrifter = strOskrifter  &"Godkendt?;Godkendt af;Faktureret?;"
				end if
				
				strOskrifter = strOskrifter  &"Kommentar;"
				
				
				objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
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
				


%>	
  
	
<%	
end if%>






	



<%if print <> "j" then%>

  <!--pagehelp-->

                <%

                itop = 84
                ileft = 735
                iwdt = 120
                ihgt = 0
                ibtop = 40 
                ibleft = 150
                ibwdt = 600
                ibhgt = 310
                iId = "pagehelp"
                ibId = "pagehelp_bread"
                call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
                %>
                       
			                <b>Visning af de-aktiverede medarbejdere: </b><br> de-aktiverede medarbejdere medtages ved søgning på "alle".<br>
			                de-aktiverede medarbejdere kan ikke vælges fra dropdown menu.<br><br>
			                <b>Jobtyper</b><br /> Fastpris eller Lbn. timer. (vægtet medarbejder timepriser)<br>
			                Ved fastpris job, hvor aktiviteterne danner grundlag for timepris, er omsætning og timepriser (på jobvisning) 
			                angivet med et ~, da beløbet er beregnet udfra en tilnærmet timepris. (gns. på akt.). Ved udspecificering på aktiviteter er det den re-elle timepris der vises.<br />
			                <br />
			                Ved lbn. timer er det altid den realiserede medarbejder timepris der vises.<br />
			                <br />
			                <b>Key-account</b><br />
                            Ved brug af "key account" bliver der vist timer for alle medarbejder der er med (tilknyttet via projektgrupper) på de job hvor den valgte Key-account er job ansvarlig / kunde ansvarlig. 
                            <br /><br />
                            <b>Beløb</b><br />
                            Alle timepriser og beløb er vist i DKK.&nbsp;
                		
                		
		                </td>
	                </tr></table></div>
	



	<br><br><br><br><br><br><br><br><br>&nbsp;
<%end if 


if print <> "j" then

ptop = 100
pleft = 920
pwdt = 140

call eksportogprint(ptop,pleft,pwdt)
%>

        
        
        
    
      <tr>
        <td align=center>
        <a href="#" onclick="Javascript:window.open('joblog.asp?FM_orderby_medarb=<%=fordelpamedarb%>&FM_medarb_hidden=<%=thisMiduse%>&media=export&datointerval=<%=strDag&"/"&strMrd&"/"&strAar & " - " & strDag_slut&"/"&strMrd_slut&"/"&strAar_slut%>&rdir=<%=rdir %>&FM_kunde=<%=kundeid %>&menu=stat&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=thisMiduse%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_kundejobans_ell_alle=<%=visKundejobans%>&FM_jobsog=<%=jobSogVal%>&FM_akttype=<%=vartyper%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
        </td><td><a href="#" onclick="Javascript:window.open('joblog.asp?FM_orderby_medarb=<%=fordelpamedarb%>&FM_medarb_hidden=<%=thisMiduse%>&media=export&datointerval=<%=strDag&"/"&strMrd&"/"&strAar & " - " & strDag_slut&"/"&strMrd_slut&"/"&strAar_slut%>&rdir=<%=rdir %>&FM_kunde=<%=kundeid %>&menu=stat&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=thisMiduse%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_kundejobans_ell_alle=<%=visKundejobans%>&FM_jobsog=<%=jobSogVal%>&FM_akttype=<%=vartyper%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport</a></td>
       </tr>
    <tr>
   <td align=center><a href="joblog.asp?FM_orderby_medarb=<%=fordelpamedarb%>&FM_medarb_hidden=<%=thisMiduse%>&FM_jobsog=<%=request("FM_jobsog")%>&FM_komprimeret=<%=komprimer%>&rdir=<%=rdir %>&FM_kunde=<%=kundeid %>&menu=stat&print=j&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=thisMiduse%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_akttype=<%=vartyper%>" target="_blank"  class='rmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
    </td><td><a href="joblog.asp?FM_orderby_medarb=<%=fordelpamedarb%>&FM_medarb_hidden=<%=thisMiduse%>&FM_jobsog=<%=request("FM_jobsog")%>&FM_komprimeret=<%=komprimer%>&rdir=<%=rdir %>&FM_kunde=<%=kundeid %>&menu=stat&print=j&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=thisMiduse%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_akttype=<%=vartyper%>" target="_blank" class="rmenu">Print version</a></td>
   </tr>
   
	
   </table>
</div>
<%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if%>



<%end if 'v <> 0 %>


</div>
<br>
<br>
&nbsp;
<%end if 'validering %>

<!--#include file="../inc/regular/footer_inc.asp"-->