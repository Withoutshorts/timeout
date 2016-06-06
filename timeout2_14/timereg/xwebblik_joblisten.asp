<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="inc/isint_func.asp"-->


<%

'section for ajax calls
if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FM_sortOrder"
Call AjaxUpdate("job","sortorder","")
End select
Response.End
end if


func = request("func")

'****************************'
'*** Opdaterer job liste ****'
'****************************'

if func = "opdaterjobliste" then


ujid = split(request("FM_jobid"), ",")

ustaar = split(request("FM_start_aar"))
ustmd = split(request("FM_start_mrd"))
ustdag = split(request("FM_start_dag"))

uslaar = split(request("FM_slut_aar"))
uslmd = split(request("FM_slut_mrd"))
usldag = split(request("FM_slut_dag"))

'uudgifter = split(request("FM_udgifterpajob"), ",")
urest = split(request("FM_restestimat"), ",")
utimer_proc = split(request("FM_stade_tim_proc"), ",")
urisiko = split(request("FM_risiko"), ",")
ustatus = split(request("FM_jobstatus"), ",")

ukomm = split(trim(request("FM_kommentar")), ", #, ")


for u = 0 to UBOUND(ujid)
	'Response.write uudgifter(u) & "<br>"
	
	ustdato = replace(ustaar(u) & "/" & ustmd(u) & "/" & ustdag(u), ",", "")
	usldato = replace(uslaar(u) & "/" & uslmd(u) & "/" & usldag(u), ",", "")
	
	
			
			
    call erDetInt(trim(urisiko(u))) 
    if isInt > 0 OR instr(trim(urisiko(u)), ".") <> 0 then
    	

    	%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
    	
	    errortype = 123
	    call showError(errortype)

    isInt = 0
    Response.end

    else
    
	strSQLjobupd = "UPDATE job SET jobstartdato = '"& ustdato &"', "_
	&" jobslutdato = '"& usldato &"', "_
	&" restestimat = "& urest(u) &", stade_tim_proc = "& utimer_proc(u) &", risiko = "& trim(urisiko(u)) &", "_
	&" jobstatus = "& ustatus(u) &", kommentar = '"& replace(ukomm(u), "'", "''") &"' WHERE id = " & ujid(u)
	oConn.execute(strSQLjobupd)
	
	'Response.write strSQLjobupd & "<br>"
	
	end if
	
next

'Response.Write "her"
'Response.flush
Response.Redirect "webblik_joblisten.asp"

end if


thisfile = "webblik_joblisten.asp"

if len(request("FM_parent")) <> 0 then
parent = request("FM_parent")
else
parent = 0
end if

if len(request("oldone")) <> 0 then
oldones = request("oldone")
else
oldones = 0
end if

if len(request("lastid")) <> 0 then
lastId = request("lastid")
else
lastId = 0
end if

print = request("print")



if len(request("realfakpertot")) <> 0 then
realfakpertot = request("realfakpertot")
Response.Cookies("webblik")("realfakpertot") = realfakpertot
else
    if request.Cookies("webblik")("realfakpertot") <> "" then
    realfakpertot = request.Cookies("webblik")("realfakpertot")
    else
    realfakpertot = 0
    end if
end if


if realfakpertot <> 0 then
realfakpertot1CHK = "CHECKED"
realfakpertot0CHK = ""
realfakpertotTxt = "Kun i periode"
else
realfakpertot1CHK = ""
realfakpertot0CHK = "CHECKED"
realfakpertotTxt = "Total, uanset valgt periode"
end if 


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	%>
	
	

	<%if print <> "j" then 
	
	dtop = "132"
	dleft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script src="inc/webblik_jav.js"></script>
	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(2)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call webbliktopmenu()
	%>
	</div>
	<%else %>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	dtop = "20"
	dleft = "20" 
	
	end if %>
	
	
	
	<%
	'*******************************************************
	'**** Igangværende Job *********************************
	'*******************************************************
	%>
	
	
	<div id="joblisten" style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px; visibility:visible; display:;">
	
	<% 
	
	oimg = "ikon_joblisten_48.png"
	oleft = 0
	otop = 0
	owdt = 600
	oskrift = "Igangværende Job"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	%>
	
	
	

	
	
	<%call filterheader(0,0,900,pTxt)%>
	
	
	<table width=100% cellpadding=0 cellspacing=0 border=0>
	<form method="post" action="webblik_joblisten.asp?FM_usedatokri=1">
	<tr>
	<td valign=top>
	
	
	<% 
	
	  
		if len(request("FM_kunde")) <> 0 then
		    
		    if len(trim(request("jobnr_sog"))) <> 0 then
			jobnr_sog = trim(request("jobnr_sog"))
			jobKri = "AND (j.jobnr = '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%') "
			show_jobnr_sog = jobnr_sog
			else
			jobnr_sog = "-1"
			jobKri = ""
			show_jobnr_sog = ""
		    end if
			
		else
		    if request.cookies("webblik")("jobnrnavn") <> "" then
			jobnr_sog = request.cookies("webblik")("jobnrnavn")
			jobKri = "AND (j.jobnr = '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%') "
			show_jobnr_sog = jobnr_sog
			else
			jobnr_sog = "-1"
			jobKri = ""
			show_jobnr_sog = ""
			end if
		end if
		
	  jobnrnavn = show_jobnr_sog 
	  response.cookies("webblik")("jobnrnavn") = jobnrnavn
	
	
	
    if len(request("FM_kunde")) <> 0 then
			
			if request("FM_kunde") = 0 then
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			else
			valgtKunde = request("FM_kunde")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			end if
			
	else
		
			if len(request.cookies("webblik")("kon")) <> 0 AND request.cookies("webblik")("kon") <> 0 then
			valgtKunde = request.cookies("webblik")("kon")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			else
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			end if
			
	end if
	
	
	response.Cookies("webblik")("kon") = valgtKunde
    
    %>
    <b>Kontakter:</b>
    
    <%if print <> "j" then %>
    <br> <select name="FM_kunde" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
		<%
		end if
		
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(valgtKunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				
				if print <> "j" then%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		</select>
		<%end if %>
		
		
    <%if print = "j" then %>
    &nbsp;<%=vlgtKunde %><br />
    <%end if %>
    
    
    
    
    <%if level <= 2 OR level = 6 then
        
                if len(request("FM_start_aar")) <> 0 then
        
                    if len(trim(request("FM_ignorer_pg"))) <> 0 then
	                IgnrProjGrp = 1
                    strPgrpSQLkri = ""
	                else
	                IgnrProjGrp = 0
		            call hentbgrppamedarb(session("mid"))
		            end if
	             
	             
	             else  
	                
	                
	                
	                    if request.cookies("webblik")("ignorer_pg") <> "" then
	                    IgnrProjGrp = request.cookies("webblik")("ignorer_pg")
	                        if cint(IgnrProjGrp) = 1 then 
    	                        strPgrpSQLkri = ""
	                        else
    	                        call hentbgrppamedarb(session("mid"))
	                        end if
                	    
	                    else
	                    IgnrProjGrp = 0
		                call hentbgrppamedarb(session("mid"))
		                end if
	               
	               end if
                	
	                
        
		if cint(IgnrProjGrp) = 1 then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if
		
		
		
		
		
		%>
		<br /><br><b>Projektgrupper:</b><br />
		<%
		
		
		if print <> "j" then%>
		<input type="checkbox" name="FM_ignorer_pg" value="1" <%=chkThis%>> Ignorer tilknytning til job via dine projektgrupper.
		<%else 
		    
		    if cint(IgnrProjGrp) = 1 then%>
		    - Ignoreret (vis alle).
		    <%else %>
		    - Viser kun de job du er tilknyttet via dine projektgrupper.
		    <%end if %>
		<%end if %>
    
    <%
    else
    IgnrProjGrp = 0
    call hentbgrppamedarb(session("mid"))
    end if
    
    response.cookies("webblik")("ignorer_pg") = IgnrProjGrp
    %>
    
    <br />  <br />
    
    
      <%
        if len(trim(request("FM_medarb_jobans"))) <> 0 then
        medarb_jobans = request("FM_medarb_jobans")
        else
                if request.cookies("webblik")("jk_ans") <> 0 then
                medarb_jobans = request.cookies("webblik")("jk_ans")
                else
                medarb_jobans = 0
                end if
        end if
        %>
         
        
        <%  
        
        if len(trim(request("job_kans"))) = 0 then
                
                if request.cookies("webblik")("jk_ans_filter") <> 0 then
                    
                    if request.cookies("webblik")("jk_ans_filter") = "1" then
                    job_kans1CHK = "CHECKED"
		            job_kans2CHK = ""
		            job_kans = "1"
		            else
		            job_kans1CHK = ""
		            job_kans2CHK = "CHECKED"
		            job_kans = "2"
		            end if
		        
                else
                    job_kans1CHK = "CHECKED"
		            job_kans2CHK = ""
		            job_kans = "1"
                end if
                
        else
        
            if request("job_kans") = "1" then
		        job_kans1CHK = "CHECKED"
		        job_kans2CHK = ""
		        job_kans = "1"
		    else
		        job_kans1CHK = ""
		        job_kans2CHK = "CHECKED"
		        job_kans = "2"
		    end if
		
		end if
    
    %>
    
    <b>Job / kundeansvarlig -  vis kun job hvor: </b><br />
           
           <%if print = "j" then%>
           - 
           <%end if %>
           
       
        
            <%if print <> "j" then%>
            <select name="FM_medarb_jobans" id="FM_medarb_jobans" style="font-size:9px; width:203px;">
            <option value="0">Alle - ignorer jobansv.</option>
            <%end if
            
            mNavn = "Alle (job / kunde ansv. ignoreret)"
            
             strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
             oRec.open strSQL, oConn, 3
             while not oRec.EOF
                
                if cint(medarb_jobans) = oRec("mid") then
                selThis = "SELECTED"
                        
                        '** medarb navn til print **
                        mNavn = oRec("mnavn") & " ("& oRec("mnr") &") "
                       
                        if len(trim(oRec("init"))) <> 0 then 
                        mNavn = mNavn & " - "& oRec("init") 
                        end if 
                        
                else
                selThis = ""
                end if
                
                if print <> "j" then%>
                <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
                 <%if len(trim(oRec("init"))) <> 0 then %>
                  - <%=oRec("init") %>
                  <%end if %>
                  </option>
                 <%
                end if
             
             oRec.movenext
             wend
             oRec.close
             %>
             
         <%if print <> "j" then%>   
        </select>
         <br />er 
        <input name="job_kans" id="job_kans" value="1" type="radio" <%=job_kans1CHK%> /> jobansvarlig  <input name="job_kans" id="Radio1" value="2" type="radio" <%=job_kans2CHK%> /> kundeansvarlig 
        
        <%else %>
     
       
            
            <%=mNavn %> er
            <%if job_kans = 1 then %>
            jobansvarlig.
            <%else %>
            kundeansvarlig.
            <%end if %>
        
        <%end if 
        
        response.cookies("webblik")("jk_ans") = medarb_jobans
        response.cookies("webblik")("jk_ans_filter") = job_kans
        
        %>
    
  
    
     <br /><br /> 
    
     <% if print <> "j" then %>
      <b>Søg på jobnr. eller jobnavn:</b>&nbsp;<br />
		    <input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=jobnrnavn%>" style="width:258px;">
		<%
		
		else %>
		<b>jobnr./jobnavn:</b><br />
		<%=jobnrnavn%>
		<%end if %>
		
	<br /><br />
	<b>Realiserede timer og faktureret beløb:<br /></b>	
	
	<% if print <> "j" then %>
        <input id="Radio2" name="realfakpertot" type="radio" value="0" <%=realfakpertot0CHK %> /> Vis <b>total</b> på viste job uanset valgt periode<br />
        <input id="Radio3" name="realfakpertot" type="radio" value="1" <%=realfakpertot1CHK %> /> Vis kun timer og beløb fra <b>den valgte periode</b>.
	
	<%else %>
	- <%=realfakpertotTxt%>
	<%end if%>
		<br /> &nbsp; 
	
	
	</td>
	
	<td valign=top style="padding-left:20px;"><b>Periode:</b><br />
	
	    <%if print = "j" then
	    dontshowDD = "1"
	    end if%>
	   
		<!--#include file="inc/weekselector_s.asp"-->
		
		<%if print = "j" then%>
		<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>
	    <%end if %>
	    
		
		
		
		
		
		<br><br />
	
	<b>Vis kun job med start eller slutdato i det valgte datointerval?</b><br />
    <%
    
    
    
    if len(request("st_sl_dato")) <> 0 then
    usedatoKri = request("st_sl_dato")  
    else
        if len(request.Cookies("webblik")("datokri")) <> 0 then
        usedatoKri = request.Cookies("webblik")("datokri")
        else
        usedatoKri = 1
        end if  
    end if
    
    response.Cookies("webblik")("datokri") = usedatoKri
    
    
    
   
    select case usedatoKri 
    case 2
    st_sl_DatoChk0 = ""
    st_sl_DatoChk1 = ""
    st_sl_DatoChk2 = "CHECKED"
    stsldatoSQLKri = "" 'jobstartdato DESC
    printVal = "- Periode ignoreret"
    
    case 1
    st_sl_DatoChk0 = ""
    st_sl_DatoChk1 = "CHECKED"
    st_sl_DatoChk2 = ""
    stsldatoSQLKri = "jobslutdato"
    printVal = "- Startdato"        
 
    case else
    st_sl_DatoChk0 = "CHECKED"
    st_sl_DatoChk1 = ""
    st_sl_DatoChk2 = ""
    stsldatoSQLKri = "jobstartdato"
    printVal = "- Slutdato"
    end select 
    
   
     
    %>
    
    <%if print <> "j" then %>
    <input id="st_sl_dato" name="st_sl_dato" type="radio" value="0" <%=st_sl_DatoChk0 %>/> Vis kun dem der har en <b>startdato</b> i det valgte datointerval.<br />
	<input id="st_sl_dato" name="st_sl_dato" type="radio" value="1" <%=st_sl_DatoChk1 %>/>  Vis kun dem der har en <b>slutdato</b> i det valgte datointerval.<br />
	<input id="st_sl_dato" name="st_sl_dato" type="radio" value="2" <%=st_sl_DatoChk2 %>/>  Ignorer periode. (Vis alle)<br />
	
	<!--<input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="checkbox" value="1" <%=stForDatoCHK%> />&nbsp;Sorter efter <u>forventet</u> slutdato.<br />-->
	<%else 
	    
	   
	    Response.Write printVal
	   
	    
	end if %>
	
	
	<%
	
		if len(request("FM_sorter")) <> 0 then
		sorter = request("FM_sorter")
		        select case sorter
		        case 1
		        orderBY = "risiko"
		        case 2
		        orderBY = "jobstartdato"
		        case else
		        orderBY = "jobslutdato"
		        end select
		else
		    if request.cookies("webblik")("prioitet") <> "" then
		    sorter = request.cookies("webblik")("prioitet")
		        select case sorter
		        case 1
		        orderBY = "risiko"
		        case 2
		        orderBY = "jobstartdato"
		        case else
		        orderBY = "jobslutdato"
		        end select
		    else
		    sorter = 0
		    orderBY = "jobslutdato"
		    end if
		    
		    
		end if 
		
		
		response.cookies("webblik")("prioitet") = sorter
		
		
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		
		
		prioTxt2 = "Periode - Startdato"
		prioTxt0 = "Periode - Slutdato" 
		prioTxt1 = "Priotet" 
		
		
		select case cint(sorter)
		case 1
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		vlgPrioTxt = prioTxt1
		
		case 2
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		vlgPrioTxt = prioTxt2
	
		
		case else
		stCHK0 = "SELECTED"
		stCHK1 = ""
		stCHK2 = ""
		
		vlgPrioTxt = prioTxt0
		
		end select
		
		
		%>
	    
	    
	    <br />
	    <br />
		<b>Sorter liste efter Periode / Prioitet:</b><br />
		
		<%if print <> "j" then %>
		
		<select name="FM_sorter" id="FM_sorter" style="background-color:#ffffe1; font-size:9px; font-family:arial; width:190px;" onchange="submit()">
		<option value="1" <%=stCHK1%>>Prioitet (drag'n drop mode)</option>
		<option value="0" <%=stCHK0%>>Slutdato</option>
		<option value="2" <%=stCHK2%>>Startdato</option>
		</select>
		
		<%else %>
		
		- <%=vlgPrioTxt %>
		
		<%end if %>
		
		
		<br /><br />
		<b>Jobstatus, vis:</b><br />
		<%if print <> "j" then %>
        <input id="Checkbox1" type="checkbox" checked disabled />
        <%else %>
        -
        <%end if %> Aktive job&nbsp;
		
		
		
		
		<%
        		
	   
	    if len(request("FM_start_aar")) <> 0 then
	        
	        if len(request("FM_status1")) <> 0 then
	        stat1 = 1
	        stCHK1 = "CHECKED"
	        else
	        stat1 = 0
	        stCHK1 = ""
	        end if
	        
	        if len(request("FM_status2")) <> 0 then
	        stat2 = 1
	        stCHK2 = "CHECKED"
	        else
	        stat2 = 0
	        stCHK2 = ""
	        end if
	        
	   
	    else
	        
	        if request.cookies("webblik")("status1") <> "" then
	                
	                stat1 = request.cookies("webblik")("status1")
	                if cint(stat1) = 1 then
	                stCHK1 = "CHECKED"
	                else
	                stCHK1 = ""
	                end if
	                
	        else
	        stat1 = 0
	        stCHK1 = ""
	        end if
	        
	        if request.cookies("webblik")("status2") <> "" then
	                
	                stat2 = request.cookies("webblik")("status2")
	                if cint(stat2) = 1 then
	                stCHK2 = "CHECKED"
	                else
	                stCHK2 = ""
	                end if
	                
	       
	        else
	        stat2 = 0
	        stCHK2 = ""
	        end if
	        
	    end if
	    
	    response.cookies("webblik")("status1") = stat1
	    response.cookies("webblik")("status2") = stat2
	    
	    
	    
	
	    
	    if print <> "j" then%>
	    <input type="checkbox" name="FM_status1" value="1" <%=stCHK1%>>Passive job&nbsp;&nbsp;
	     <input type="checkbox" name="FM_status2" value="2" <%=stCHK2%>>Lukkede job
	    <%else %>
	    
	        <%if stat1 = "1" then%>
	        - Passive job<br />
	        <%end if %>
	        
	         <%if stat2 = "1" then%>
	        - Lukkede job<br />
	        <%end if %>
	    
	    <%end if %>
	    
	    
	     <br />
   
   <% '** Skulte job **'
   if len(request("st_sl_dato")) <> 0 then
   
   if request("visskjulte") <> 0 then
   visskjulte = 1
   CHKskj = "CHECKED"
   else
   visskjulte = 0
   CHKskj = ""
   end if
   
   else
   
   
        if request.cookies("webblik")("visskj") <> "1" then
        visskjulte = 0
        CHKskj = ""
        else
        visskjulte = 1
        CHKskj = "CHECKED"
        end if
        
   end if
   
   response.cookies("webblik")("visskj") = visskjulte
    %>
    
    <%if print <> "j" then %>
        <input id="visskjulte" name="visskjulte" type="checkbox" <%=CHKskj %> value="1" /> Vis skjulte job (prioitet = -1)
    <%else 
        if visskjulte = 1 then%>
        - Viser skjulte job 
        <%else %>
        - Viser ikke skjulte job 
        <%end if %>
    <%end if %>
	    
	    <%
	    response.cookies("webblik").expires = date + 60
	    %>
		
		
    </td>
	<td valign=bottom>
	<%if print <> "j" then%>
	<input type="submit" value="Vis joblisten >>">
	<%end if%>
	</td></tr>
	</form>
	</table>
	
	<!-- filter div -->
	</td></tr></table>
	</div>
	
	<%if print <> "j" then %>
	            <br /><br /><table cellpadding=0 cellspacing=0 border=0><tr><td bgcolor="#ffffff" style="padding:0px 0px 0px 5px; border:1px #8cAAe6 solid; border-right:0px;">
	            <a href='jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik' class='vmenu' alt="Opret nyt job" target="_top">Opret nyt job</a>
                    </td><td bgcolor="#ffffff" style="padding:4px 5px 0px 5px; border:1px #8cAAe6 solid; border-left:0px;">
                    <a href='jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik' class='vmenu' alt="Opret nyt job" target="_top"><img src="../ill/add2.png" border="0" /></a>
                    </td></tr></table>
	            
	            <%end if %>
	
	<br />
	<br />
	<%if sorter = "1" then
	tbID = "incidentlist"
	else
	tbID = ""
	end if  %>
	<table cellspacing=0 cellpadding=0 border=0 width=99% id="<%=tbID%>">
	<form method="post" action="webblik_joblisten.asp?func=opdaterjobliste">
    
    <!--
    <input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="hidden" value="<%=st_sl_dato_forventet%>"/>
	-->
	
	<% 
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	%>
	
    <%if print <> "j" then %>
	<tr>
	<td colspan=5 style="padding:5px;">
	   <%
	    if sorter = "1" then
	    uWdt = 300
	    uTxt = "<b>Sortering:</b><br>Træk i et job for at prioitere rækkefølgen. <br>Prioitets-angivelsen skifter først ved gen-indlæs."
	    call infoUnisport(uWdt, uTxt) 
	    else
	    Response.Write "&nbsp;"
	    end if%>
	    
	</td>
	<td colspan=10 align=right style="padding:20px 10px 10px 0px;"><input type="submit" value="Opdater liste"></td>
    </tr>
	<%end if %>
	
	
	
	<tr height=30 bgcolor="#5582d2">
		<td class=alt style="padding:3px;" valign=bottom><b>Kontakt</b><br />
		Job<br />
		Jobbesk.<br />
		Job. ansv. & ejer</td>
		<td class=alt style="padding:3px;" valign=bottom><b>Periode</b></td>
		<td class=alt style="padding:3px;" valign=bottom><b>Status</b><br />Prioitet<br />
		(-1 = skjul)</td>
		
		<td class=alt style="padding:3px;" valign=bottom><b>Forecast</b><br /> (allokeret tim. <br />også de-aktive medarb.)</td>
		<td class=alt style="padding:3px;" valign=bottom><b>Forkalk. tim.</b><br />på job</td>
		<td class=alt style="padding:3px;" valign=bottom><b>Real. timer</b></td>
		
		<td class=alt style="padding:3px;" valign=bottom>
		<b>Realiseret %</b></td>
		<td class=alt style="padding:3px; width:150px;" valign=bottom><b>Rest. estimat</b><br />
		(Stade indmelding i %)<br />
		Forventet samlet tidsforbr.<br />
		Balance og Afvigelse %</td>
		<td class=alt style="padding:3px 10px 3px 3px;" valign=bottom align=right></td>
		<td class=alt style="padding:3px;" valign=bottom>Bal.</td>
		<td class=alt style="padding:3px; width:105px;" valign=bottom align=right><b>Brutto oms. DKK</b><br />
		Udgft. u.lev. + Intern<br />
		</td>
		<td class=alt style="padding:3px; width:35px;" valign=bottom align=right>DB %</td>
		<td class=alt style="padding:3px; width:105px;" valign=bottom align=right>
		<b>Faktureret beløb DKK</b><br />
		~ Gns. faktisk timepris<br />
		Balance brutto oms. / Faktureret</td>
		<td class=alt style="padding:5px;" colspan=2 valign=bottom><b>Kommentar</b><br />
		Maks: 200 kar.</td>
	</tr>
	<%
	
	
	if cint(stat1) = 1 then
	statKri = "jobstatus = 1 OR jobstatus = 2"
	else
	statKri = "jobstatus = 1"
	end if
	
	if cint(stat2) = 1 then
	statKri = statKri & " OR jobstatus = 0 "
	else
	statKri = statKri
	end if
	
	if medarb_jobans <> 0 then
	    if job_kans = 1 then
	    kansKri = ""
	    jobansKri = " AND (j.jobans1 = "& medarb_jobans &" OR j.jobans2 = "& medarb_jobans &") "
	    else
	    kansKri = " AND (kundeans1 = "& medarb_jobans &" OR kundeans2 = "& medarb_jobans &")"
	    jobansKri = ""
	    end if
	else
	kansKri = ""
	jobansKri = ""
	end if
	
	if cint(visskjulte) = 1 then
	visskjulteKri = ""
	else
	visskjulteKri = " AND risiko >= 0"
	end if
	
	jids = 0
	
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobstatus, "_
	&" jobans2, kkundenr, jobslutdato, jobstartdato, j.beskrivelse, jobans1, jobans2, m.mnavn, m2.mnavn AS mnavn2,"_
	&" ikkebudgettimer, budgettimer, jobtpris, sum(r.timer) AS restimer, stade_tim_proc, "_
	&" risiko, udgifter, rekvnr, forventetslut, restestimat, jobstatus, j.kommentar, s.navn AS aftnavn, jo_dbproc "_
	&" FROM job j LEFT JOIN kunder ON (kid = jobknr) "_
	&" LEFT JOIN medarbejdere m ON (m.mid = jobans1)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
	&" LEFT JOIN ressourcer_md r ON (r.jobid = j.id)"_
	&" LEFT JOIN serviceaft s ON (s.id = serviceaft)"_
	&" WHERE fakturerbart = 1 "& visskjulteKri &" AND ("& statKri &") "& jobKri &" "
	
	if cint(usedatoKri) <> 2 then
	strSQL = strSQL &" AND "&  stsldatoSQLKri &" BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
	end if
	
	strSQL = strSQL & jobansKri & strPgrpSQLkri &" AND "& sqlKundeKri &" "& kansKri &""_
	&" GROUP BY j.id ORDER BY "& orderBY &", kkundenavn" 
	
	
	'Response.write strSQL
	'Response.flush
	
	c = 0
	budgetIalt = 0
	budgettimerIalt = 0
	restimerIalt = 0
	lastMonth = 0
	tilfakturering = 0
	udgifterTot = 0
	timerforbrugtIalt = 0
	lastsortid = 0
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	
	faktureret = 0
	timerTildelt = 0
	timerforbrugt = 0
	totalforbrugt = 0
	totalBalance = 0
	jobbudget = 0
	
	if cint(oRec("jobstatus")) = 2 then '(passiv)
	stname = "<font class=megetlillesort>(passivt)</font> "
	else
	stname = ""
	end if
	
	bg = right(c, 1)
	select case bg
	case 0,2,4,6,8
	bgthis = "#ffffff"
	case else
	bgthis = "#eff3ff"
	end select
	
	
	
	'** Rettigheder til at redigere **
	if level = 1 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			end if
	end if
	%>
	
	<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=oRec("id")%>">
	
	<!--<input id="hd_jobid" value="<%=jobid%>" type="hidden" />-->
	
	
	<% 
	'if isDate(oRec("forventetslut")) = true then
	'forvSlutDato = oRec("forventetslut") 
	'else
	'forvSlutDato = oRec("jobslutdato")
	'end if
	
	
	btop = 1
	
	select case sorter 
	case 2
	datoFilterUse = oRec("jobstartdato")
	case 1
	datoFilterUse = oRec("jobslutdato")
	case else
	datoFilterUse = oRec("jobslutdato")
	end select
	
	
	if cint(sorter) = 0 OR cint(sorter) = 2 then
	
	    if datepart("m", datoFilterUse, 2, 3) <> lastMonth OR datepart("yyyy", datoFilterUse, 2, 3) <> lastYear then
	        if c <> 0 then%>
	         <!--<tr>
		    <td height="20" colspan="15">&nbsp;</td>
	        </tr>-->
	        <%end if %>
	    <tr bgcolor="#ffffe1">
		    <td style="padding:5px 0px 5px 10px; border:1px orange solid;" colspan=15><b><%=ucase(monthname(datepart("m", datoFilterUse, 2, 3))) &" "& year(datoFilterUse)%></b></td>
	    </tr>
	    <%
	    btop = 0
	    end if
	
	lastMonth = datepart("m", datoFilterUse, 2, 3) 
	lastYear = datepart("yyyy", datoFilterUse, 2, 3) 
	
	end if
	%>
	
	
	
	<tr bgcolor="<%=bgthis%>">
		<td valign=top style="padding:6px 5px 3px 3px; width:230px; border-top:<%=btop%>px #cccccc dashed;">
		<b><%=oRec("kkundenavn")%></b>
            &nbsp;(<%=oRec("kkundenr") %>)<br>
		   
		    <%if print <> "j" then%>
		        <%if editok = 1 then%>
		        <a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=webblik" class=vmenu><%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>)&nbsp;</a>
		        <a href="jobs.asp?menu=job&func=slet&id=<%=oRec("id")%>&jobnr_sog=a&filt=aaben&fm_kunde_sog=0&rdir=webblik" class=red>[x]</a>
		        <%else%>
		        <b><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</b>
		        <%end if%>
		    <%else %>
		    <%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)
		    <%end if %>
		
		<%
		
		if len(trim(oRec("beskrivelse"))) <> 0 then
		htmlparseCSV(oRec("beskrivelse")) 
		strBesk = htmlparseCSVtxt 
		else
		strBesk = ""
		end if
		%>
		<font class="megetlillesort"><i>
		<%if len(strBesk) > 100 then %>    
		<br /><%=left(strBesk, 100) %>..
		<%else %>
		<br /><%=strBesk %>
		<%end if %>
		</i></font>
		
		<font class="megetlillesilver"><i><%=oRec("mnavn")%>
		<%if len(trim(oRec("mnavn2"))) <> 0 then %>
		, <%=oRec("mnavn2")%>
		<%end if %>
		</i></font>
		
		<% 
		if len(trim(oRec("rekvnr"))) <> 0 then
	    Response.Write "<br>Rekvnr.:  "& oRec("rekvnr") 
	    end if
		%>
		
		<%
		
		if len(trim(oRec("aftnavn"))) <> 0 then
	    Response.Write "<br>Aft.: "& oRec("aftnavn") 
	    end if
		
		%>
		<input type="hidden" name="rowId" value="<%=oRec("id")%>" />
		</td>
		
		
		<td valign=top style="padding:6px 3px 3px 3px; width:230px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid;" class=lille>
		
		<!--
		<=formatdatetime(oRec("jobstartdato"), 2)%> til <font color=darkred><%=formatdatetime(oRec("jobslutdato"), 2)%></font><br>
		-->
					<%if print <> "j" then%>
					<select name="FM_start_dag" id="Select2" style="font-size:9px; font-family:arial;">
										<option value="<%=datepart("d",oRec("jobstartdato"), 2, 2)%>"><%=datepart("d",oRec("jobstartdato"), 2, 2)%></option> 
									   	<option value="1">1</option>
									   	<option value="2">2</option>
									   	<option value="3">3</option>
									   	<option value="4">4</option>
									   	<option value="5">5</option>
									   	<option value="6">6</option>
									   	<option value="7">7</option>
									   	<option value="8">8</option>
									   	<option value="9">9</option>
									   	<option value="10">10</option>
									   	<option value="11">11</option>
									   	<option value="12">12</option>
									   	<option value="13">13</option>
									   	<option value="14">14</option>
									   	<option value="15">15</option>
									   	<option value="16">16</option>
									   	<option value="17">17</option>
									   	<option value="18">18</option>
									   	<option value="19">19</option>
									   	<option value="20">20</option>
									   	<option value="21">21</option>
									   	<option value="22">22</option>
									   	<option value="23">23</option>
									   	<option value="24">24</option>
									   	<option value="25">25</option>
									   	<option value="26">26</option>
									   	<option value="27">27</option>
									   	<option value="28">28</option>
									   	<option value="29">29</option>
									   	<option value="30">30</option>
										<option value="31">31</option></select>
					
					<select name="FM_start_mrd" id="Select3" style="font-size:9px; font-family:arial;">
					<option value="<%=datepart("m",oRec("jobstartdato"), 2, 2)%>"><%=left(monthname(datepart("m",oRec("jobstartdato"), 2, 2)), 3)%></option>
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
					
					
					<select name="FM_start_aar" id="Select4" style="font-size:9px; font-family:arial;">
					<option value="<%=datepart("yyyy",oRec("jobstartdato"), 2, 2)%>"><%=datepart("yyyy",oRec("jobstartdato"), 2, 2)%></option>
					<%for x = -5 to 10
		            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		            <option value="<%=right(useY, 2)%>"><%=useY%></option>
		            <%next %>
					</select>&nbsp;&nbsp;
							
							
				<%else %>
				
				<%=datepart("d",oRec("jobstartdato"), 2, 2)%>.<%=datepart("m",oRec("jobstartdato"), 2, 2)%>.<%=datepart("yyyy",oRec("jobstartdato"), 2, 2)%>
				
				<%end if %>
					
					<%if print <> "j" then%>
					<br /><select name="FM_slut_dag" id="FM_slut_dag" style="font-size:9px; font-family:arial;">
										<option value="<%=datepart("d",oRec("jobslutdato"), 2, 2)%>"><%=datepart("d",oRec("jobslutdato"), 2, 2)%></option> 
									   	<option value="1">1</option>
									   	<option value="2">2</option>
									   	<option value="3">3</option>
									   	<option value="4">4</option>
									   	<option value="5">5</option>
									   	<option value="6">6</option>
									   	<option value="7">7</option>
									   	<option value="8">8</option>
									   	<option value="9">9</option>
									   	<option value="10">10</option>
									   	<option value="11">11</option>
									   	<option value="12">12</option>
									   	<option value="13">13</option>
									   	<option value="14">14</option>
									   	<option value="15">15</option>
									   	<option value="16">16</option>
									   	<option value="17">17</option>
									   	<option value="18">18</option>
									   	<option value="19">19</option>
									   	<option value="20">20</option>
									   	<option value="21">21</option>
									   	<option value="22">22</option>
									   	<option value="23">23</option>
									   	<option value="24">24</option>
									   	<option value="25">25</option>
									   	<option value="26">26</option>
									   	<option value="27">27</option>
									   	<option value="28">28</option>
									   	<option value="29">29</option>
									   	<option value="30">30</option>
										<option value="31">31</option></select>
					
					<select name="FM_slut_mrd" id="FM_slut_mrd" style="font-size:9px; font-family:arial;">
					<option value="<%=datepart("m",oRec("jobslutdato"), 2, 2)%>"><%=left(monthname(datepart("m",oRec("jobslutdato"), 2, 2)), 3)%></option>
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
					
					
					<select name="FM_slut_aar" id="FM_slut_aar" style="font-size:9px; font-family:arial;">
					<option value="<%=datepart("yyyy",oRec("jobslutdato"), 2, 2)%>"><%=datepart("yyyy",oRec("jobslutdato"), 2, 2)%></option>
					<%for x = -5 to 10
		            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		            <option value="<%=right(useY, 2)%>"><%=useY%></option>
		            <%next %></select>&nbsp;&nbsp;
					
					<br />
					<a href="webblik_joblisten21.asp?nomenu=1&FM_kunde=<%=oRec("jobknr") %>&jobnr_sog=<%=oRec("jobnr") %>" target="_blank" class=rmenu>Gantt >></a>
							
				<%else %>
				
				- <%=datepart("d",oRec("jobslutdato"), 2, 2)%>.<%=datepart("m",oRec("jobslutdato"), 2, 2)%>.<%=datepart("yyyy",oRec("jobslutdato"), 2, 2)%>
				
				<%end if %>
		</td>
		
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:0px #cccccc solid; width:80px;" class=lille>
		<%
		
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		
		
		select case oRec("jobstatus")
		case 1
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		stName = "Aktiv"
		case 2
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		stName = "Passiv"
		case 0
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stName = "Lukket"
		end select
		
		
		if print <> "j" then%>
		<select name="FM_jobstatus" id="Select1" style="font-size:9px; font-family:arial;">
		<option value="0" <%=stCHK0%>>Lukket</option>
		<option value="1" <%=stCHK1%>>Aktiv</option>
		<option value="2" <%=stCHK2%>>Passiv</option>
		</select>
		<%else %>
		
		<%=stName%>
		
		<%end if %>
		<br />
		<%
		
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		selbgcol = ""
		
		select case oRec("risiko")
		case 1
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		stName = "Lav"
		case 2
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		stCHK3 = ""
		stCHK4 = ""
		stName = "Mellem"
		case 3
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = "SELECTED"
		stCHK4 = ""
		stName = "Høj"
		case 4
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = "SELECTED"
		stName = "Skjult"
		case else
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		stCHK3 = ""
		stCHK4 = ""
		stName = ""
		end select
		
		if cdbl(lastsortid) >= oRec("risiko") AND oRec("risiko") >= 0 then
		thissortID = lastsortid + 1
		else
		thissortID = oRec("risiko")
		end if
		
		if oRec("risiko") >= 0 then
		lastsortid = thissortID
		end if
		
		if print <> "j" then%>
		
		<!--
		<select name="FM_risiko" id="FM_risiko" style="font-size:9px; font-family:arial;">
		<option value="0" <%=stCHK0%>>Ej angivet</option>
		<option value="1" <%=stCHK1%>>Lav</option>
		<option value="2" <%=stCHK2%>>Mellem</option>
		<option value="3" <%=stCHK3%>>Høj</option>
		<option value="4" <%=stCHK4%>>Skjul</option>
		</select>
		-->
		 <%'if sorter = "1" then 
		 'prioFMType = "hidden"
		 'else
		 prioFMType = "text"
		 'end if%>
         <input name="FM_risiko" id="FM_risiko" type="<%=prioFMType %>" value="<%=thissortID%>" style="width:30px; font-size:8px; font-family:arial;" />
         
		<%else %>
		<%=thissortID%>
        <%end if %>
		</td>
		
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #CCCCCC solid; border-left:1px #CCCCCC solid;" align=right><%=formatnumber(oRec("restimer"), 2)%></td>
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid;" align=right>
		
		<%
		timerTildelt = oRec("ikkebudgettimer") + oRec("budgettimer")
		if len(timerTildelt) <> 0 then
		timerTildelt = timerTildelt
		else
		timerTildelt = 0
		end if
		
		Response.write "<b>"& formatnumber(timerTildelt, 2) & "</b>"
		%>
		  <input id="FM_forkalk_<%=oRec("id") %>" type="hidden" value="<%=timerTildelt%>" />
		
		</td>
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid;" align=right>
		<%
		call akttyper2009(2)
		
		'*** Timer forbrugt ***
		strSQL2 = "SELECT sum(t.timer) AS timerforbrugt FROM timer t WHERE t.tjobnr = "& oRec("jobnr") &" AND ("& aty_sql_realhours &") "
		
		if realfakpertot <> 0 then
		strSQL2 = strSQL2 &" AND tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		end if
		
		strSQL2 = strSQL2 &" GROUP BY t.tjobnr"
		
		'Response.Write strSQL2
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		timerforbrugt = oRec2("timerforbrugt")
		end if
		oRec2.close
		
		if len(timerforbrugt) <> 0 then
		timerforbrugt = timerforbrugt
		else
		timerforbrugt = 0
		end if
		
		timerforbrugtIalt = timerforbrugtIalt + timerforbrugt
		
		Response.write formatnumber(timerforbrugt, 2)
		%>
            <input id="FM_timerreal_<%=oRec("id")%>" type="hidden" value="<%=timerforbrugt%>" />
		</td>
		
		
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid;" class=lille><%
		
		
		
		
		if timerTildelt <> 0 then
		projektcomplt = ((timerforbrugt/timerTildelt)*100)
	    else
	    projektcomplt = 100
	    end if
	    
	    if projektcomplt >= 100 then
		fcol = "crimson"
		showprojektcomplt = projektcomplt
		projektcomplt = 100
		else
		fcol = "#DCF5BD"
		showprojektcomplt = projektcomplt
		end if
		
		
		if timerforbrugt > 0 then%>
		
		<div style="width:100px; border:1px #CCCCCC solid; height:20px;">
		<span style="width:<%=cint(left(projektcomplt, 3))%>px; background-color:<%=fcol%>; height:20px; padding:3px 0px 3px 3px;">
		<%if showprojektcomplt > 0 then%>
		<%=formatpercent(showprojektcomplt/100, 0)%>
		<%end if%>
		</span>
		</div>
		<%
		
		else
		
		Response.Write "&nbsp;"
		
		end if
		
		%>
		
		</td>
		
		<td valign=top style="padding:6px 3px 3px 3px;; border-top:<%=btop%>px #cccccc dashed;" class=lille nowrap>
		<%
		if len(oRec("restestimat")) <> 0 then
		restestimat = oRec("restestimat")
		else
		restestimat = 0
		end if
		%>
		
		
		<%
		
		select case oRec("stade_tim_proc") '0 = rest i timer, 1 = i proc
		case 0
		totalforbrugt = (timerforbrugt + restestimat)
		totalBalance = (timerTildelt - totalforbrugt)
		case 1
		    
		    if timerforbrugt <> 0 OR restestimat <> 0 then
		    totalforbrugt = timerforbrugt * (100/restestimat) 
		    else
		    totalforbrugt = 0
		    end if '- (restestimat/timerTildelt) * 100)
		
		totalBalance = timerTildelt - totalforbrugt '(timerforbrugt + (timerTildelt - (restestimat/timerTildelt) * 100))
		end select
		
		
		
		if totalforbrugt <> 0 then
		
		if totalBalance >= 0 then
		bgst = "#DCF5BD"
		else
		bgst = "crimson"
		end if
		
		else
		
		bgst = "#cccccc"
		
		end if %>
		
		
		<%if print <> "j" then%>
		<input type="text" name="FM_restestimat" id="FM_restestimat_<%=oRec("id")%>" class="rest" value="<%=restestimat%>" style="font-size:9px; font-family:arial; width:50px;">
		
		<%select case oRec("stade_tim_proc")
		case 0
		stade_timproc_SEL0 = "SELECTED"
		stade_timproc_SEL1 = ""
		case 1
		stade_timproc_SEL0 = ""
		stade_timproc_SEL1 = "SELECTED"
		end select %>
		<select class="stade" name="FM_stade_tim_proc" id="FM_stade_tim_proc_<%=oRec("id")%>" style="font-size:9px; font-family:arial; width:100px;">
		    <option value=0 <%=stade_timproc_SEL0 %>>timer til rest.</option>
		    <option value=1 <%=stade_timproc_SEL1 %>>% afsluttet</option>
		</select>
		<%else 
		
		select case oRec("stade_tim_proc")
		case 0
		stade_timproc_txt = "timer til rest."
		case 1
		stade_timproc_txt = "% afsluttet"
		end select
		%>
		<%=restestimat & " " & stade_timproc_txt%> 
		<%end if %>
		
		
		
		<div id="totalfb_<%=oRec("id")%>">Forv. samlet tidsforbr.: <%=formatnumber(totalforbrugt, 2)%></div>
		
		
		<table border=0 cellpadding=0 cellspacing=0>
		<tr>
		<td class=lille>
		<div id="totalbal_<%=oRec("id")%>">Balance: <%=formatnumber(totalBalance, 2)%></div>
		
		</td>
		<td class=lille>&nbsp;</td>
		<td class=lille>
		<div id="totalafv_<%=oRec("id")%>">  
		
		
		
		
		<%if timerTildelt <> 0 and totalforbrugt <> 0 then
		
		if timerTildelt < totalforbrugt then
		prc = 100 - (timerTildelt/totalforbrugt) * 100
		else
		prc = 100 - (totalforbrugt/timerTildelt)* 100
		end if
		    
		    Response.write "&nbsp;(Afv: "& formatnumber(prc, 0) & "%) "
		
		else
	        
	        Response.write "&nbsp;" '& formatnumber(100, 0) & "%"
		
		end if
		%>
		
		</div>
		</td></tr></table>
		
		
		<div id="totalstade_<%=oRec("id")%>"> 
		
		<%if cint(oRec("stade_tim_proc")) = 0 then 'rest estimat angivet i timer
		
		      if totalforbrugt <> 0 then
		
		        if restestimat = 0 then
		        'restestimat = timerTildelt '** = 100%
        		
		        stade = 100
        		
		        else
        		
        		stade = (timerforbrugt/totalforbrugt) * 100
		            
		            'if timerTildelt < restestimat then
		            'stade = 100 - (timerTildelt/restestimat * 100)
		            'else
		            'stade = 100 - (restestimat/timerTildelt * 100)
		            'end if
        		
		        end if
		    
		    else
		    
		    stade = 0
		    
		    end if
		    
		    Response.write "Stade: ~ "& formatnumber(stade, 0) & "% afsluttet"
		
		else
	        
	        Response.write "Stade: ~ " & formatnumber(restestimat, 0) & "% afsluttet"
		
		end if
		%>
		
		</div> 
		
		
		</td>
		
		<td valign=top style="padding:6px 10px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid;" class=lille align=right>
		&nbsp;
		</td>
		
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed; border-right:1px #cccccc solid;" class=lille align=right>
		<div id="divindikator_<%=oRec("id")%>" style="width:15px; border:1px #CCCCCC solid; background-color:<%=bgst%>; height:15px;"><img src="ill/blank.gif" border="0" /></div>
	    </td>
		
		
		<td valign=top class=lille style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed;" align=right>
		<b><%
		if len(oRec("jobTpris")) <> 0 then
		jobbudget = oRec("jobTpris")
		else
		jobbudget = 0
		end if
		Response.write formatnumber(jobbudget, 2) & " DKK"
		%></b>
		
		<br />
		<%
		if len(oRec("udgifter")) <> 0 then
		udgifterpajob = oRec("udgifter")
		else
		udgifterpajob = 0
		end if
		
		%>
		<%=formatnumber(udgifterpajob, 2)%> DKK 
		
		<%
		udgifterTot = udgifterTot + udgifterpajob
		%>
		
		</td>
		
		
		
		
		<td valign=top align=right bgcolor="#ffdfdf" style="padding:6px 3px 3px 3px; border-left:1px #cccccc solid; border-right:1px #cccccc solid; border-top:<%=btop%>px #cccccc dashed;">
		
		
		<%=formatnumber(oRec("jo_dbproc"),0) %> %
		
		</td>
		
		<%
		
		'*** Faktureret **'
		faktureret = 0
		aktbelFak = 0
		'timerFak = 0
		
		'*** Faktureret ***'
		strSQL2 = "SELECT sum(f.beloeb * (f.kurs/100)) AS faktureret, sum(f.timer) AS timer, "_
		&" SUM(fd.aktpris* (f.kurs/100)) AS aktbel "_
		&" FROM fakturaer AS f "_
		&" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid)"_
		&" WHERE jobid = "& oRec("id") &" AND faktype = 0 AND shadowcopy = 0 "_
		&" AND fd.enhedsang <> 3 "
		
		if realfakpertot <> 0 then
		strSQL2 = strSQL2 &" AND fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		end if
		
		strSQL2 = strSQL2 &" GROUP BY jobid"
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		faktureret = oRec2("faktureret")
		'timerFak = oRec2("timer")
		aktbelFak = oRec2("aktbel")
		end if
		oRec2.close
		
		'*** Kredit ***'
		aktbelKre = 0
		
		strSQL2 = "SELECT sum(f.beloeb * (f.kurs/100)) AS faktureret, sum(f.timer) AS timer, "_
		&" SUM(fd.aktpris* (f.kurs/100)) AS aktbel "_
		&" FROM fakturaer AS f "_
		&" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid)"_
		&" WHERE jobid = "& oRec("id") &" AND faktype = 1 AND shadowcopy = 0 "_
		&" AND fd.enhedsang <> 3 "
		
		if realfakpertot <> 0 then
		strSQL2 = strSQL2 &" AND fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		end if
		
		strSQL2 = strSQL2 &" GROUP BY jobid"
		
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		faktureretKre = oRec2("faktureret")
		'timerFakKre = oRec2("timer")
		aktbelKre = oRec2("aktbel")
		end if
		oRec2.close
		
		
		if faktureret <> 0 then
		faktureret = faktureret - (faktureretKre)
		else
		faktureret = 0
		end if
		
		
		aktBel = aktbelFak - (aktbelKre)
		'if timerFak <> 0 then
		'timerFak = timerFak - (timerFakKre)
		'else
		'timerFak = 0
		'end if
		
		if timerforbrugt <> 0 then
		gnstimepris = aktBel/timerforbrugt
		else
		gnstimepris = 0
		end if
		
		faktotbel = faktotbel + (faktureret)
		'faktottim = faktottim + (timerFak)
		
		%>
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed;" class=lille align=right>
		<b><%=formatnumber(faktureret) & " DKK" %><br /></b>
		<%=formatnumber(gnstimepris) %> DKK<br />
		
		<%
		bal = (faktureret - jobbudget)
		
		if bal < 0 then
		fcol = "red"
		else
		fcol = "green"
		end if
		
		Response.write "<font color="& fcol &">" & formatnumber(bal, 2) & " DKK" 
		
		tilfakturering = tilfakturering + (bal)
		%>
		</td>
		<td valign=top colspan=2 class=lille style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc dashed;">
		<%if print <> "j" then %>
		<textarea id="FM_kommentar" name="FM_kommentar" cols="20" rows="3" style="width:133px; font-size:9px; font-family:Arial;" ><%=oRec("kommentar") %></textarea>	
            <input id="FM_kommentar" name="FM_kommentar" value="#" type="hidden" />	
            <%else %>
            <%=oRec("kommentar") %>
            <%end if %>		
		</td>
		</tr>
	
	<%
	
	budgetIalt = budgetIalt + oRec("jobtpris")
	budgettimerIalt = budgettimerIalt + (oRec("budgettimer") + oRec("ikkebudgettimer"))
	
	if len(oRec("restimer")) <> 0 then
	restimerIalt = restimerIalt + oRec("restimer")
	else
	restimerIalt = restimerIalt
	end if
	
	'** Til Export fil ***'
	jids = jids & "," & oRec("id")
	
	Response.flush
	
	c = c + 1
	oRec.movenext
	wend
	oRec.close 
	
	%>
	
	 <input id="FM_kommentar" name="FM_kommentar" value="xc" type="hidden" />	
	
	
	<tr bgcolor="#ffdfdf">
	<td colspan=3 style="border-top:1px #cccccc solid;">&nbsp;</td>
	<td valign=top align=right bgcolor="#c4c4c4" style="border-top:1px #999999 solid; border-left:1px #999999 solid; border-right:1px #999999 solid; padding:6px 3px 3px 3px;" nowrap><b><%=formatnumber(restimerIalt, 2)%> t.</b> </td>
	<td valign=top align=right style="border-top:1px #cccccc solid; border-right:1px #cccccc solid; padding:6px 3px 3px 3px;" nowrap><b><%=formatnumber(budgettimerIalt, 2)%> t.</b></td>
	<td valign=top align=right style="border-top:1px #cccccc solid; border-right:0px #cccccc solid; padding:6px 3px 3px 3px;" nowrap><b><%=formatnumber(timerforbrugtIalt, 2) %> t.</b></b></td>
    <td colspan=4 align=right style="border-top:1px #cccccc solid; border-right:1px #cccccc solid;">
        &nbsp;</td>
	<td align=right valign=top style="border-top:1px #cccccc solid; padding:6px 8px 3px 3px;"><b><%=formatnumber(budgetIalt, 2)%> DKK</b><br />
	<%=formatnumber(udgifterTot, 2) %> DKK</td>
	
	<td align=center style="border-left:1px #cccccc solid; border-right:1px #cccccc solid; border-top:1px #cccccc solid; padding:6px 0px 3px 0px;">
	
	&nbsp;
	
	</td>
	<td align=right valign=top style="border-top:1px #cccccc solid;padding:6px 3px 3px 3px;" nowrap>
	<b><%=formatnumber(faktotbel, 2) %> DKK<br /></b>
	<%=formatnumber(tilfakturering, 2) %> DKK</td>
	<td align=right style="border-top:1px #cccccc solid;padding:6px 3px 3px 3px;" colspan=2>
        &nbsp;</td>
	
	</tr>
	
	
	<%if print <> "j" then %>
	<tr><td colspan=15 align=right style="padding:20px 10px 0px 0px;"><input type="submit" value="Opdater liste"></td></tr>
	<%end if %>
	
	
	</form>
	</table>
	
	
	
	
	
	
	<%if print <> "j" then
	
	itop = 20
	ileft = 0
	iwdt = 400
	
	call sideinfo(itop,ileft,iwdt)
	
	%>
	
	<b>Prioitet</b><br />
	For at være sikker på at opnå end bedre prioitet end den forrige på listen skriv da 0,9 for at blive vist bedre end aktuelle nr 1 på listen.
	<br /><br />
	<b>Skjulte job</b><br />
	sæt prioitet = -1
	</td>
	</tr>
	</table>
	</div>
	
	<%
	
	            ptop = 72
                pleft = 915
                pwdt = 130

                call eksportogprint(ptop,pleft, pwdt)
                %>
                
                
                    
                     <tr>
                    <td align=center>
                   <a href="job_eksport.asp?jids=<%=jids%>" class=rmenu target="_blank">
                   &nbsp;<img src="../ill/export1.png" border=0 alt="Eksport" /></a>
                </td><td><a href="job_eksport.asp?jids=<%=jids%>" class=rmenu target="_blank">.csv fil eksport</a></td>
                   </tr>
                    
                    <tr>
                    <td align=center>
                   <a href="webblik_joblisten.asp?print=j" target="_blank">
                   &nbsp;<img src="../ill/printer3.png" border=0 alt="Print version" /></a>
                </td><td><a href="webblik_joblisten.asp?print=j" target="_blank" class=rmenu>Print version</a></td>
                   </tr>
                  
                    </table>
                      </div>
                   
                   
           
                <%else%>

                <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %>
                <%end if%>
	
	
	
	<br /><br />
        &nbsp;
	
	</div>
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->