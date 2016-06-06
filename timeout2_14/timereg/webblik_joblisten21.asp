<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
func = request("func")

'*** Opdaterer job liste ****
if func = "ajaxupdate" then

Function AjaxFormatDates(datestring,interval)
if (Not IsNull(datestring)) AND (Not IsNull(interval)) then
if InStr(datestring,"/") then
splitslashdate = split(datestring,"/")
datestring1 = splitslashdate(2)&"-"&splitslashdate(1)&"-"&splitslashdate(0)
end if
datestring1 = DateAdd("d",interval,datestring)
datestringsplit = split(datestring1,"-")
AjaxFormatDates = datestringsplit(2)&"/"&datestringsplit(1)&"/"&datestringsplit(0)
end if
End Function

strAjaxDataType = Request.Form("datatype")
strAjaxStartdate = Request.Form("startdate")
strAjaxEnddate = Request.Form("enddate")
strAjaxCnr = Request.Form("cnr")
strAjaxJid = Request.Form("jid")
strAjaxDateDiff = Request.Form("DateDiff")
strAjaxId = Request.Form("id")
'strAjaxCalcStartdate = AjaxFormatDates(strAjaxStartdate, strAjaxDateDiff)
'strAjaxCalcEnddate = AjaxFormatDates(strAjaxEnddate, strAjaxDateDiff)
returnstring = "cnr= " & strAjaxCnr & " jid=" & strAjaxJid & " datediff= " & strAjaxDateDiff & " id=" & strAjaxId & " startdate =" & strAjaxStartdate & " enddate =" &strAjaxEnddate&" formattedstartdate ="'&AjaxFormatDates(strAjaxStartdate)
'returnstring = returnstring & " added datediffstartdate = " & strAjaxCalcStartdate &" added datediffenddate = " & strAjaxCalcEnddate
response.Write returnstring

select case strAjaxDataType

case "milepole"

	strSQLajaxmpupd = "UPDATE milepale SET dato = '" &strAjaxStartdate& "' WHERE id = " & strAjaxId
	Response.Write ("/n " & strSQLajaxmpupd)
	oConn.execute(strSQLajaxmpupd)

case "prekon_0"

	strSQLajaxmpupd = "UPDATE job SET preconditions_met = 1 WHERE id = " & strAjaxJid
	oConn.execute(strSQLajaxmpupd)

   

case "prekon_1"

	strSQLajaxmpupd = "UPDATE job SET preconditions_met = 2 WHERE id = " & strAjaxJid
    oConn.execute(strSQLajaxmpupd)

    
case "prekon_2"

	strSQLajaxmpupd = "UPDATE job SET preconditions_met = 0 WHERE id = " & strAjaxJid
	oConn.execute(strSQLajaxmpupd)

    
case "prioitet"
    
  	
	'isInt = 0
    prioVal = 0
    if len(trim(request.form("prioval"))) <> 0 then
        prioVal = request.form("prioval")

        'call erDetInt(prioVal) 
        'if isInt = 0 then

        strSQLajaxmpupd = "UPDATE job SET risiko = "& prioVal &" WHERE id = " & strAjaxJid
	    oConn.execute(strSQLajaxmpupd)

        'end if

    end if
  

case "job"
'ustdato = replace(ustaar(u) & "/" & ustmd(u) & "/" & ustdag(u), ",", "")

	strSQLajaxjobupd = "UPDATE job SET jobstartdato = '"& strAjaxStartdate &"', "_
	&" jobslutdato = '"& strAjaxEnddate &"' WHERE id = " & strAjaxJid
	Response.Write ("/n " & strSQLajaxjobupd)
	oConn.execute(strSQLajaxjobupd)
	
	strSQL2 = "SELECT aktstartdato, aktslutdato, id FROM aktiviteter WHERE job = " & strAjaxJid

	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF

	strSQLajaxaktupd = "UPDATE aktiviteter SET aktstartdato = '" & AjaxFormatDates(oRec2("aktstartdato"), strAjaxDateDiff) & "', "_
	&" aktslutdato = '" & AjaxFormatDates(oRec2("aktslutdato"), strAjaxDateDiff) & "' WHERE id = " & oRec2("id")
	oConn.execute(strSQLajaxaktupd)
	response.Write strSQLajaxaktupd
	oRec2.movenext
	wend
	oRec2.close
	
	strSQL3 = "SELECT dato, id FROM milepale WHERE jid = " & strAjaxJid
	oRec3.open strSQL3, oConn, 3
	while not oRec3.EOF
	
	strSQLajaxmpupd = "UPDATE milepale SET dato = '" & AjaxFormatDates(oRec3("dato"), strAjaxDateDiff) & "' WHERE id = " & oRec3("id")
	Response.Write ("/n " & strSQLajaxmpupd)
	oConn.execute(strSQLajaxmpupd)
	
	oRec3.movenext
	wend
	oRec3.close

	
	response.Write "updated job, activities and milepoles"

	
case "activity"

	strSQLajaxaktupd = "UPDATE aktiviteter SET aktstartdato = '" & strAjaxStartdate & "', "_
	&" aktslutdato = '" & strAjaxEnddate & "' WHERE id = " & strAjaxId
	oConn.execute(strSQLajaxaktupd)
	Response.Write ("/n " & strSQLajaxaktupd)
	response.Write "updated job"

end select

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
		<!--include file="../inc/regular/topmenu_inc.asp"-->

		<%
    	
	    errortype = 123
	    call showError(errortype)

    isInt = 0
    Response.end

    else
    
	strSQLjobupd = "UPDATE job SET jobstartdato = '"& ustdato &"', "_
	&" jobslutdato = '"& usldato &"', "_
	&" restestimat = "& urest(u) &", risiko = "& trim(urisiko(u)) &", "_
	&" jobstatus = "& ustatus(u) &", kommentar = '"& replace(ukomm(u), "'", "''") &"' WHERE id = " & ujid(u)
	oConn.execute(strSQLjobupd)
	
	'Response.write strSQLjobupd & "<br>"
	
	end if
	
next

'Response.Write "her"
'Response.flush
Response.Redirect "webblik_joblisten21.asp"

end if




if func = "gemliste" then
    
    
    
    jobnrs = request("jobnrs")

    if len(trim(request("listeid"))) <> 0 then
    listeid = request("listeid")
    else
    listeid = 0
    end if

    if len(trim(request("listenavn"))) <> 0 then
    listenavn = request("listenavn")
    else
    listenavn = ""
    end if

    dtMidt = request("dtmidt")


    tstamp = year(now) &"/"& month(now) &"/"& day(now)

    if listenavn = "" then

    
    	%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<!--include file="../inc/regular/topmenu_inc.asp"-->

		<%
    	
	    errortype = 161
	    call showError(errortype)



    response.end
    end if

    if request("slet_liste") <> "1" then

    if request("gem_liste") <> "1" AND cint(listeid) <> 0 then
    strSQLg = "UPDATE gantts SET jobnrs = '"& jobnrs &"', navn = '"& listenavn &"', dato = '"& tstamp &"', editor = '"& session("user") &"', datomidtpkt = '"& dtMidt &"' WHERE id = "& listeid
    else
    strSQLg = "INSERT INTO gantts (jobnrs, navn, dato, editor, datomidtpkt) VALUES ('"& jobnrs &"', '"& listenavn &"', '"& tstamp &"', '"& session("user") &"', '"& dtMidt &"')"
    end if

    else
        
        
        strSQLg = "DELETE FROM gantts WHERE id = " & listeid

    end if

    'Response.write strSQLg
    'Response.flush

    oConn.execute(strSQLg)

   

    Response.redirect "webblik_joblisten21.asp"
    
end if












thisfile = "webblik_joblisten21.asp"

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

if len(trim(request("nomenu"))) <> 0 and request("nomenu") <> "0" then
nomenu = 1
else
nomenu = 0
end if

if len(trim(request("FM_bruglistedato"))) <> 0 then
bruglistedato = request("FM_bruglistedato")
else
bruglistedato = 0
end if

'Response.write "bruglistedato:" & bruglistedato & "<br>" 


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
	    if nomenu <> 1 then
	        dtop = "62"
	        dleft = "90" 
        	
	        %>
	        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	        <!--include file="../inc/regular/topmenu_inc.asp"-->

            <!--
	        <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	        <%call tsamainmenu(2)%>
	        </div>
	        <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	        <%
	        call webbliktopmenu()
	        %>
	        </div>
                -->
                

	     <%call menu_2014()
             
             
            else 
	     
	        dtop = "20"
	        dleft = "20" 
	     
	     %>
	      <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	      <%end if %>
	<%else %>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<!--
	<!--include file="../inc/regular/header_hvd_inc.asp"
	<script src="../inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>
	     
    <link href="../inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
    -->
	<%
	
	dtop = "20"
	dleft = "20" 
	
	end if %>



	<script type="text/javascript" src="../inc/jquery/jquery.jobliste.js"></script>
	<script type="text/javascript">

          
            	 $(window).load(function() {


         PopulateGraphFrame();

		    <% if (print <> "j") then%>
            //AddJobs();
			AddJobs({searchId : <%="'"&request("jobnr_sog")&"'"%>});
			<%else %>
			AddJobs({print : true});
			<%end if %>

            $("#todayMarker").height($("#JobGraph").height());
            $("#tilLabel").hide();
			$("#FM_enddate").css("display", "none");




              /// Pre-conditions ///
                        $(".prekon_0").bind('mouseover', function () {
                            $(this).css('cursor', 'pointer');
                        });
          

                        $(".prekon_0").bind('click', function () {
                        var thisid = this.id

                        var idlngt = thisid.length
                        var idtrim = thisid.slice(7, idlngt)

       
                        prekon_0(idtrim);

                                   
                        });

                        


                        $(".prekon_1").bind('mouseover', function () {
                            $(this).css('cursor', 'pointer');
                        });
          

                        $(".prekon_1").bind('click', function () {
                        var thisid = this.id

                        var idlngt = thisid.length
                        var idtrim = thisid.slice(7, idlngt)

       
                        prekon_1(idtrim);

                                        

                                      

                        });


                         
                        $(".prekon_2").bind('mouseover', function () {
                            $(this).css('cursor', 'pointer');
                        });
          

                        $(".prekon_2").bind('click', function () {
                        var thisid = this.id

                        var idlngt = thisid.length
                        var idtrim = thisid.slice(7, idlngt)

       
                        prekon_2(idtrim);
                                

                                   

                        });



                        $(".opd_prio").bind('mouseover', function () {
                            $(this).css('cursor', 'pointer');
                        });
          

                        $(".opd_prio").bind('click', function () {
                        var thisid = this.id

                        var idlngt = thisid.length
                        var idtrim = thisid.slice(9, idlngt)

                        //alert("her" +idtrim)

                        opdater_prioitet(idtrim);
                                

                                   

                        });


                       


             


                 


         });





             $(document).ready(function() {


             $("#jobnr_sog").keyup(function () {
                 $("#FM_gantt_liste").val('0');
             });
            


             $("#FM_gantt_liste").change(function () {
                 $("#FM_bruglistedato").val('1');

                 $("#filterform").submit();
             });
            

            
            });



	




        
     

	</script>
	
	<%
	'*******************************************************
	'**** Igangværende Job *********************************
	'*******************************************************
	%>
	
	
	<div id="joblisten" style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px; visibility:visible; display:;">
	
	<% 
	
	'oimg = "ikon_gantt_48.png"
	'oleft = 0
	'otop = 0
	'owdt = 600
	oskrift = "Jobplanner Gantt"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
    if print = "j" then
    %>
    Zoom: [CTRL] +/-
    <%
    end if


    '***************************************************************************'
	'** variable ***
    '***************************************************************************'

    public jobKri
    function jobSogSql(jobnr_sog) 
         
         dim jobKriStrArr
         
         'Response.write jobnr_sog
         'Response.flush

         jobSogFundet = 0   
            
            if instr(jobnr_sog, ",") <> 0 then
                
                jobKri = " AND (j.jobnr = '0' "
                jobKriStrArr = split(jobnr_sog, ",")
                for j = 0 to UBOUND(jobKriStrArr)
                jobKri = jobKri & " OR j.jobnr = '" & trim(jobKriStrArr(j)) & "'"
                next
                jobKri = jobKri & ")"

                jobSogFundet = 1

            end if


             if instr(jobnr_sog, "-") <> 0 then
                
                
                jobKriStr_instr = instr(jobnr_sog, "-")
                jobKriStr_len = len(jobnr_sog)

                jobKriStr_left = left(jobnr_sog, jobKriStr_instr-1)
                jobKriStr_right = mid(jobnr_sog, jobKriStr_instr+1, jobKriStr_len)

               
                jobKri = " AND (j.jobnr BETWEEN "& jobKriStr_left &" AND "& jobKriStr_right &")"

                jobSogFundet = 2

            end if


            if jobSogFundet = 0 then
            jobKri = "AND (j.jobnr LIKE '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%' OR k.kkundenavn LIKE '"& jobnr_sog &"%') "
            end if

    end function


    if len(request("FM_kunde")) <> 0 then
		    
		    if len(trim(request("jobnr_sog"))) <> 0 then
			jobnr_sog = trim(request("jobnr_sog"))
            call jobSogSql(jobnr_sog)
            show_jobnr_sog = jobnr_sog
            else
			jobnr_sog = "-1"
			jobKri = ""
			show_jobnr_sog = ""
		    end if
			
		else
		
		    if request.cookies("webblik")("jobnrnavn") <> "" then
			jobnr_sog = request.cookies("webblik")("jobnrnavn")
			call jobSogSql(jobnr_sog)
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



    if level <= 2 OR level = 6 then
        
        if len(request("FM_start_aar")) <> 0 then
        
        if len(request("FM_ignorer_pg")) <> 0 then
        IgnrProjGrp = 1
        else
        IgnrProjGrp = 0
        end if
        
		if cint(IgnrProjGrp) = 1 then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if
		
		else
		    
		    if request.cookies("webblik")("prjgrp") <> "" then
		    IgnrProjGrp = request.cookies("webblik")("prjgrp")
		        if cint(IgnrProjGrp) = 1 then
		        chkThis = "CHECKED"
		        else
		        chkThis = ""
		        end if
		     else
		     IgnrProjGrp = 0
		     chkThis = ""
		     end if
		
		end if
		
		response.cookies("webblik")("prjgrp") = IgnrProjGrp
 
   end if


        if len(trim(request("FM_medarb_jobans"))) <> 0 then
        medarb_jobans = request("FM_medarb_jobans")
        else
                if request.cookies("webblik")("jk_ans") <> 0 then
                medarb_jobans = request.cookies("webblik")("jk_ans")
                else
                medarb_jobans = 0
                end if
        end if
       
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
		

          response.cookies("webblik")("jk_ans") = medarb_jobans
        response.cookies("webblik")("jk_ans_filter") = job_kans



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



    	if len(request("FM_sorter")) <> 0 then
		sorter = request("FM_sorter")
		        select case sorter
		        case 1
		        orderBY = "risiko"
		        case 2
		        orderBY = "jobstartdato DESC"
		        case else
		        orderBY = "jobslutdato DESC"
		        end select
		else
		    if request.cookies("webblik")("prioitet") <> "" then
		    sorter = request.cookies("webblik")("prioitet")
		        select case sorter
		        case 1
		        orderBY = "risiko"
		        case 2
		        orderBY = "jobstartdato DESC"
	            case else
		        orderBY = "jobslutdato DESC"
		        end select
		    else
		    sorter = 0
		    orderBY = "jobslutdato"
		    end if
		    
		    
		end if 
		
		
		response.cookies("webblik")("prioitet") = sorter
		        

                '*** Aktiviteter sorter ***'
		        select case sorter
		        case 1
		        AorderBY = "a.fase, a.sortorder, a.navn"
		        case 2
		        AorderBY = "a.aktstartdato, a.fase, a.sortorder, a.navn"
		        case else
		        AorderBY = "a.aktslutdato, a.fase, a.sortorder, a.navn"
		        end select



        

        if len(trim(request("interval"))) <> 0 then
        interval = request("interval")
        else
            if request.cookies("webblik")("interval") <> "" then
            interval = request.cookies("webblik")("interval")
            else
            interval = 0
            end if
        end if
         

        if interval <> 0 then
        interval1CHK = "CHECKED"
        interval0CHK = ""
        else
        interval1CHK = ""
        interval0CHK = "CHECKED"
        end if
		
		response.cookies("webblik")("interval") = interval



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
	    

        if len(trim(request("FM_gantt_liste"))) <> 0 AND request("FM_gantt_liste") <> 0 then
        gantt_liste = request("FM_gantt_liste")
            

            strSQLglst = "SELECT jobnrs FROM gantts WHERE id = " & gantt_liste
            oRec4.open strSQLglst, oConn, 3
            while not oRec4.EOF 
            
            jobnrsUse = oRec4("jobnrs")
            
            oRec4.movenext
            wend  
            oRec4.close


            call jobSogSql(jobnrsUse)

        else
        gantt_liste = 0
        end if

        listenavn = ""

	    response.cookies("webblik")("status1") = stat1
	    response.cookies("webblik")("status2") = stat2
	    


         response.cookies("webblik").expires = date + 60


    '***********************************************************************************'
    '**********************************************************************************'
  
    %>
    <form id="filterform" method="post" action="webblik_joblisten21.asp?FM_usedatokri=1&nomenu=<%=nomenu%>">
    <%
    if print = "j" then

     
	    dontshowDD = "1"
	    %>
	    <!--#include file="inc/weekselector_jq.asp"-->
		
		<input type="hidden" name="FM_startdate" id="FM_startdate" value="<%=strDag&"/"&strMrd&"/"&strAar%>" />
	    <input type="radio" name="interval" id="interval0" value="0" <%=interval0CHK %> style="visibility:hidden;" />
		<input type="radio" name="interval" id="interval1" value="1" <%=interval1CHK %> style="visibility:hidden;" />
    <%
    dontshowDD = ""
    end if





    '****************************************************************************************
    '*************************** Filter          ********************************************
    '****************************************************************************************
    if print <> "j" then
      'call filterheader(0,0,800,pTxt)
        call filterheader_2013(40,0,800,oskrift) 
    %>
	
	
	<table width=100% cellpadding=0 cellspacing=0 border=0>
	
	<tr>
	<td valign=top>
	
	

    <b>Kontakter:</b>
    
  
    <br> <select name="FM_kunde" size="1" style="width:325px;" onChange="submit();">
		<option value="0">Alle</option>
		<%
		
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
				
			%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				
				%>
		</select>
		
		
		
  
		<br /><br><b>Projektgrupper:</b><br />
		<%
		
		
	%>
		<input type="checkbox" name="FM_ignorer_pg" value="1" <%=chkThis%>> Ignorer tilknytning til job via dine projektgrupper.
		
    
   
    <br />  <br />
    
    
    
    <b>Job / kundeansvarlig -  vis kun job hvor: </b><br />
           
     
           
            <select name="FM_medarb_jobans" id="FM_medarb_jobans" style="width:325px;">
            <option value="0">Alle - ignorer jobansv.</option>
            <% 
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
                
           %>
                <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
                 <%if len(trim(oRec("init"))) <> 0 then %>
                  - <%=oRec("init") %>
                  <%end if %>
                  </option>
                 <%
                
             
             oRec.movenext
             wend
             oRec.close
             %>
             
         
        </select>
         <br />er 
        <input name="job_kans" id="job_kans" value="1" type="radio" <%=job_kans1CHK%> /> jobansvarlig (1-2)  <input name="job_kans" id="Radio1" value="2" type="radio" <%=job_kans2CHK%> /> kundeansvarlig 
        
    
      
   <br /><br />
   
  
    <b>Skjulte job</b><br />
  
        <input id="visskjulte" name="visskjulte" type="checkbox" <%=CHKskj %> value="1" /> Vis skjulte job (prioitet = -1)
   
  <br /><br />
  
  

    <b>Søg på jobnr., job- eller. kunde -navn:</b><br />(% wildcard, brug "," for flere job, eller "-" for interval)&nbsp;<br />
		<input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=jobnrnavn%>" style="width:325px; border:2px #6CAE1C solid;">
		
		
         <br />
         <br />
         <b>Vis gemte gannt lister:</b> (viser alle job på den gemte liste)
         <%
         strSQLg = "SELECT id, jobnrs, navn, dato, datomidtpkt FROM gantts WHERE id <> 0 ORDER BY navn" 
         
         

         %>
         <select name="FM_gantt_liste" id="FM_gantt_liste" style="width:325px;">
         <option value="0">Ingen</option>
         <%
         oRec4.open strSQLg, oConn, 3
         While not oRec4.EOF


         if cint(gantt_liste) = oRec4("id") then
         gnt_sel = "SELECTED"
         listenavn = oRec4("navn")
         listeid = oRec4("id")
         datomidtpkt = oRec4("datomidtpkt")
         else
         gnt_sel = ""
         end if 
         %> <option value="<%=oRec4("id") %>" <%=gnt_sel %>><%=oRec4("navn") %> (datomidtpunkt: <%=oRec4("datomidtpkt") %>)</option><%
         
         oRec4.movenext
         wend
         oRec4.close
         
         
        
         %>

        
        
         </select>
         <input type="hidden" value="0" name="FM_bruglistedato" id="FM_bruglistedato" />
		<br /><span style="color:#999999; font-size:9px;">Viser valgte job på listen og tager udgangspunkt i det gemte datomidtpunkt.</span>
	



	</td><td valign=top><b>Grafisk datomidtpunkt:</b><br />
	
	 
	  
		<!--#include file="inc/weekselector_jq.asp"-->
	
		
		
		
		
		
		
        <!--
	<b>Start eller Slutdato?</b><br />
  
    <input id="st_sl_dato" name="st_sl_dato" type="radio" value="0" <%=st_sl_DatoChk0 %>/> Vis kun dem der har en <b>startdato</b> i det valgte datointerval.<br />
	<input id="st_sl_dato" name="st_sl_dato" type="radio" value="1" <%=st_sl_DatoChk1 %>/>  Vis kun dem der har en <b>slutdato</b> i det valgte datointerval.<br />
	<input id="st_sl_dato" name="st_sl_dato" type="radio" value="2" <%=st_sl_DatoChk2 %>/>  Ignorer periode. (Vis alle)<br />
	-->

    <input id="st_sl_dato" name="st_sl_dato" type="hidden" value="2"/>

	<!--<input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="checkbox" value="1" <%=stForDatoCHK%> />&nbsp;Sorter efter <u>forventet</u> slutdato.<br />-->
	
	
	
		<br /><br /><br />

        <%sorter = request("FM_sorter")
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
		
		
		prioTxt2 = "Periode - Startdato"
		prioTxt0 = "Periode - Slutdato" 
		prioTxt1 = "Priotet" 
		
		

       

        'sorter = 1
        'response.write sorter
        'response.flush

		select case sorter
		case "1"
		stCHK0 = ""
		stCHK1 = "SELECTED"
		stCHK2 = ""
		vlgPrioTxt = prioTxt1
		
		case "2"
		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = "SELECTED"
		vlgPrioTxt = prioTxt2
	
		
		case else
		stCHK0 = "SELECTED"
		stCHK1 = ""
		stCHK2 = ""
		
		vlgPrioTxt = prioTxt0
		
		end select %>


		<b>Sorter liste efter Periode / Prioitet:</b><br />
	
		
		<select name="FM_sorter" style="width:225px;" ><!--id="FM_sorter" onchange="submit()" --> 
		<option value="0" <%=stCHK0%>>Periode - Slutdato</option>
		<option value="1" <%=stCHK1%>>Prioitet</option>
		<option value="2" <%=stCHK2%>>Periode - Startdato</option>
		</select>
		
		
		<br /><br /><br />
        <b>Visning:</b><br />
        <input type="radio" name="interval" id="interval0" value="0" <%=interval0CHK %> onclick="submit()" /> 12 måneder (md/md)<br />
		<input type="radio" name="interval" id="interval1" value="1" <%=interval1CHK %> onclick="submit()"/> Kvartal (uge/uge)<br />

        

		<br /><br />
		<b>Jobstatus, vis</b><br />
		<input type="checkbox" name="" value="1" DISABLED CHECKED>Aktive&nbsp;<br />
	    <input type="checkbox" name="FM_status1" value="1" <%=stCHK1%>>Passive / Til fakturering<br />
	     <input type="checkbox" name="FM_status2" value="2" <%=stCHK2%>>Lukkede 



	   
	  
		
    </td>
	<td valign=bottom>
	
	<input type="submit" value="Vis joblisten >>">
	
	</td></tr>
	
	</table>
	
    	<!-- filter div -->
	</td></tr></table>
	</div>
    

    <%end if 'print %>
    </form>


    




	<br />
	<br />
	
	<br />
	<br />



    <div id="JobGraph">
    	<table cellspacing=0 cellpadding=0 border=0 width=100%>
    <tr style="background-color:#FFFFFF;"><td style="height:15px;">


<ul id="JobGraphHeader">
<li class="infoField" style="width:200px;">
Kontakt (Knr.)
Jobnavn (Jobnr.)
</li>
</ul>

</td>
</tr>
<tr style="background-color:#F7F7F7;"><td style="height:30px;">

<%if cint(interval) = 1 then %>
<ul id="JobGraphHeader_weekday">
<li class="infoField" style="width:200px;">
&nbsp;
</li>
</ul>
<%end if %>

</td>
</tr>
<tr><td>


<%if print <> "j" then %>
<span id="navPrevMonth" style="float:left; cursor:hand;" class="ui-icon ui-icon-circle-arrow-w"></span><span id="navNextMonth" style="float:right; cursor:hand;" class="ui-icon ui-icon-circle-arrow-e"></span><br class="clear" />
<%end if %>
<div id="JobList">

</div>
	<center><img class="loadingIcon" src="../inc/jquery/images/ajax-loader.gif" alt="henter data"/></center>
</div>

</div>

 </td>
    </tr>
	</table>


</div><!-- jobGraph -->
    <!--
    <input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="hidden" value="<%=st_sl_dato_forventet%>"/>
	-->
	
	<% 

    Function JSFormat(jsstring)
	if not IsNull(jsstring) then
  jsstring = replace(jsstring, ",",".")
  jsstring = replace(jsstring, "'","\'")
  end if
  JSFormat = jsstring
	End Function


    if interval = 0 then
	sqlDatostartsplit = split(DateAdd("m",-6,strDag&"/"&strMrd&"/"&strAar),"-")
	sqlDatoslutsplit = split(DateAdd("m",6,strDag&"/"&strMrd&"/"&strAar),"-")
	sqlDatostart = sqlDatostartsplit(2) & "/" & sqlDatostartsplit(1) & "/" & sqlDatostartsplit(0)  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut =  sqlDatoslutsplit(2) & "/" & sqlDatoslutsplit(1) & "/" & sqlDatoslutsplit(0)'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	else
    sqlDatostartsplit = split(DateAdd("m",-2,strDag&"/"&strMrd&"/"&strAar),"-")
	sqlDatoslutsplit = split(DateAdd("m",1,strDag&"/"&strMrd&"/"&strAar),"-")
	sqlDatostart = sqlDatostartsplit(2) & "/" & sqlDatostartsplit(1) & "/" & sqlDatostartsplit(0)  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut =  sqlDatoslutsplit(2) & "/" & sqlDatoslutsplit(1) & "/" & sqlDatoslutsplit(0)'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	end if


	if request("FM_ignorer_pg") <> "1" then
		call hentbgrppamedarb(session("mid"))
	else
		strPgrpSQLkri = ""
	end if
	
	    
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

    call akttyper2009(2)
	
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobstatus, "_
	&" jobans2, kkundenr, jobslutdato, jobstartdato, j.beskrivelse, jobans1, jobans2, m.mnavn, m2.mnavn AS mnavn2, m.email, m2.email AS email2,"_
	&" ikkebudgettimer, budgettimer, jobtpris, sum(r.timer) AS restimer, "_
	&" risiko, udgifter, rekvnr, forventetslut, restestimat, jobstatus, j.kommentar, s.navn AS aftnavn, jo_dbproc, preconditions_met"_
	&" FROM job AS j LEFT JOIN kunder AS k ON (kid = jobknr) "_
	&" LEFT JOIN medarbejdere m ON (m.mid = jobans1)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
	&" LEFT JOIN ressourcer_md r ON (r.jobid = j.id)"_
	&" LEFT JOIN serviceaft s ON (s.id = serviceaft)"_
	&" WHERE fakturerbart <> -99 "& visskjulteKri &" AND ("& statKri &") "& jobKri
	
	if cint(usedatoKri) <> 2 then
	strSQL = strSQL &" AND "&  stsldatoSQLKri &" BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
	end if
	
	strSQL = strSQL & jobansKri & strPgrpSQLkri &" AND "& sqlKundeKri &" "& kansKri &""_
	&" GROUP BY j.id ORDER BY "& orderBY &", kkundenavn" 
	
	strSQL = strSQL & " LIMIT 50"
	
    'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQL
	'Response.flush
    'Response.write "<br /><br /><br /><a href=""#"" id=""showarr"">link</a>"

    jids = 0
    lastJob = 0
	jobnrs = 0

	c = 0
	budgetIalt = 0
	budgettimerIalt = 0
	restimerIalt = 0
	lastMonth = 0
	tilfakturering = 0
	udgifterTot = 0
	timerforbrugtIalt = 0
	lastsortid = 0


    call akttyper2009(2)
	
	%>
	<script type="text/javascript">
	<%

    AddJobListDBnr = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
   
    'Response.write "<hr>her<br>"& c & "<br>"
    'Response.flush

	faktureret = 0
	timerTildelt = 0
	timerforbrugt = 0
	totalforbrugt = 0
	totalBalance = 0
	jobbudget = 0
	
	
	
	'** Rettigheder til at redigere **
	if level = 1 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			end if
	end if

	btop = 1
	
	select case sorter 
	case 2
	datoFilterUse = oRec("jobstartdato")
	case 1
	datoFilterUse = oRec("jobslutdato")
	case else
	datoFilterUse = oRec("jobslutdato")
	end select
	

		
		if len(trim(oRec("aftnavn"))) <> 0 then
	    'Response.Write "<br>Aft.: "& oRec("aftnavn") 
	    end if
		
		timerTildelt = oRec("ikkebudgettimer") + oRec("budgettimer")
		if len(timerTildelt) <> 0 then
		timerTildelt = timerTildelt
		else
		timerTildelt = 0
		end if

		
        
        '*** Timer forbrugt ***


        aty_sql_realhoursT = replace(aty_sql_realhours, "a.fakturerbar", "t.tfaktim")
        timerforbrugt = 0
		strSQLtimer = "SELECT sum(t.timer) AS timerforbrugt FROM timer t WHERE t.tjobnr = '"& oRec("jobnr") &"' AND ("& aty_sql_realhoursT &") GROUP BY t.tjobnr"
        'strSQLtimerTxt = strSQLtimerTxt & "<br>::"& strSQLtimer
        'Response.flush
        oRec2.open strSQLtimer, oConn, 3
		
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
		
		'Response.write formatnumber(timerforbrugt, 2)

		if len(oRec("restestimat")) <> 0 then
		restestimat = oRec("restestimat")
		else
		restestimat = 0
		end if

		totalforbrugt = (timerforbrugt + restestimat)
		'Response.write "= <u>"&formatnumber(totalforbrugt, 2)&"</u>"
        
		totalBalance = (timerTildelt - totalforbrugt)
		
		if totalBalance < 0 then
		fcol = "red"
		else
		fcol = "green"
		end if
		
        if timerTildelt <> 0 and totalforbrugt <> 0 then
		
		if timerTildelt < totalforbrugt then
		prc = 100 - (timerTildelt/totalforbrugt) * 100
		else
		prc = 100 - (totalforbrugt/timerTildelt)* 100
		end if
		    
		    'Response.write " "& formatnumber(prc, 0) & "%"
		
		else
	        
	        'Response.write "&nbsp;" '& formatnumber(100, 0) & "%"
		
		end if
		
		
		'*** Faktureret **'
		faktureret = 0
		timerFak = 0
		
		'*** Faktureret ***'
		strSQL2 = "SELECT sum(beloeb * (kurs/100)) AS faktureret, sum(timer) AS timer FROM fakturaer WHERE jobid = "& oRec("id") &" AND faktype = 0 GROUP BY jobid"
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		faktureret = oRec2("faktureret")
		timerFak = oRec2("timer")
		end if
		oRec2.close
		
		'*** Kredit ***'
		strSQL2 = "SELECT sum(beloeb * (kurs/100)) AS faktureret, sum(timer) AS timer FROM fakturaer WHERE jobid = "& oRec("id") &" AND faktype = 1 GROUP BY jobid"
		oRec2.open strSQL2, oConn, 3
		
		if not oRec2.EOF then
		faktureretKre = oRec2("faktureret")
		timerFakKre = oRec2("timer")
		end if
		oRec2.close
		
		
		if faktureret <> 0 then
		faktureret = faktureret - (faktureretKre)
		else
		faktureret = 0
		end if
		
		if timerFak <> 0 then
		timerFak = timerFak - (timerFakKre)
		else
		timerFak = 0
		end if
		
		faktotbel = faktotbel + (faktureret)
		faktottim = faktottim + (timerFak)
		
		bal = (faktureret - jobbudget)
		
		if bal < 0 then
		fcol = "red"
		else
		fcol = "green"
		end if
		
		'Response.write "<font color="& fcol &">" & formatnumber(bal, 2) &" "& basisValISO_f8 
		
		tilfakturering = tilfakturering + (bal)
		%>	

	ArrActivities = new Array();
	ArrMilePael = new Array();
	<%
	
	
            

	aty_sql_realhoursA = replace(aty_sql_realhours, "tfaktim", "a.fakturerbar")
	
	strSQL2 = "SELECT a.navn, a.id AS aid, a.budgettimer, a.fakturerbar, sum(t.timer) AS timer, t.tdato, t.timerkom, t.offentlig, "_
	&" t.sttid, t.sltid, "_
	&" a.incidentid, a.aktstartdato, a.aktslutdato, a.beskrivelse, a.fase "_
	&" FROM aktiviteter a "_
	&" LEFT JOIN timer t ON (t.taktivitetid = a.id)"_
    &" LEFT JOIN ressourcer_md r ON (r.aktid = a.id)"_
	&" WHERE a.job = " & oRec("id") & " AND ("& aty_sql_realhoursA &") GROUP BY a.id ORDER BY "& AorderBY
	
	ActivityNr = 0
	
    'response.write "<br>XXXXXXXXXXXXXXXXXXXXXXXXXX<br>Aktiviteter:<br>"& strSQL2
   
    'Response.flush

	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF
	
	
	                '*** Ressourcetimer på aktivitet ***
	                resTimerAkt = 0
                    strSQL3 = "SELECT sum(r.timer) AS restimer FROM ressourcer_md AS r WHERE jobid = "& oRec("id") &" AND r.aktid = "& oRec2("aid") &" GROUP BY aktid"
                    oRec3.open strSQL3, oConn, 3
		
                    if not oRec3.EOF then
                    resTimerAkt = oRec3("restimer")
                    end if
                    oRec3.close

	//make realisedoutput no matter what
	if IsNull(oRec2("timer")) then
	Arealiserettimer = 0
	else
	Arealiserettimer = oRec2("timer")
	end if
	
	if IsNull(oRec2("budgettimer")) OR oRec2("budgettimer") = 0 then
    Abudgettimer = 0
    else
    Abudgettimer = oRec2("budgettimer")
	end if
	
    if isNull(oRec2("fase")) then
    lcsFase = ""
    else
    lcsFase = "<br><span style=""color:#6CAE1C; size:9px;"">fase: " & lcase(replace(oRec2("fase"), "_", " ")) & "</span>"
    end if
	%>
	ArrActivities[<%=ActivityNr%>] = { "name" : <%="'"& JSFormat(left(oRec2("navn"),30)) & lcsFase &"'"%>, "precalc" : <%="'"&JSFormat(formatnumber(Abudgettimer, 0))&"'"%>, "realised" : <%="'"&JSFormat(formatnumber(Arealiserettimer,0))&"'"%>, 'resource' : <%="'"&JSFormat(formatnumber(resTimerAkt, 0))&"'"%>, "startdate" : <%="'"&oRec2("aktstartdato")&"'"%>, "enddate" : <%="'"&oRec2("aktslutdato")&"'"%>, "id" : <%="'"&oRec2("aid")&"', 'datatype' : 'activity'"%> };
	<%		
	ActivityNR = ActivityNR + 1
	oRec2.movenext
	wend
	oRec2.close



	
	strSQL3 = "Select m.id, m.navn, m.dato, m.beskrivelse, mt.navn AS mttypenavn FROM milepale m LEFT JOIN milepale_typer mt ON (mt.id = m.type) where m.jid = "& oRec("id")
	'response.write strSQL3
	MilePaelNR = 0
	
	oRec3.open strSQL3, oConn, 3
	while not oRec3.EOF
	
	
	%>
	ArrMilePael[<%=MilePaelNR%>] = { "name" : <%="'"&JSFormat(oRec3("navn"))&"'"%>, "type" : <%="'"&JSFormat(oRec3("mttypenavn"))&"'"%>, "startdate" : <%="'"&JSFormat(oRec3("dato"))&"'"%>, "desc" : <%="'"&oRec3("beskrivelse")&"'"%>, "id" : <%="'"&oRec3("id")&"', 'datatype' : 'milepole'"%> };
	<%		
	MilePaelNR = MilePaelNR + 1
	oRec3.movenext
	wend
	oRec3.close
	


    'txt = txt & "her<br>"& c & "<br>"& oRec("jobnavn")
    'Response.flush 
    %>
    //alert(ArrActivities)
    //AddJob(ArrActivities, customer, name, id, precalc, realised, resource, startdate, enddate)
    AddJobListDB[<%=AddJobListDBnr%>] = {'ArrActivities' : ArrActivities, 'ArrMilePael' : ArrMilePael, <%="'customer' : '"&JSFormat(left(oRec("kkundenavn"), 25)) &"' ,'name':'"&JSFormat(left(oRec("jobnavn"), 25))&"','jnr': '"&oRec("jobnr")&"','jid': '"&oRec("id")&"','cnr': '"&oRec("kkundenr")&"','precalc': '"&JSFormat(formatnumber(timerTildelt,0))&"','realised': '"&JSFormat(formatnumber(timerforbrugt,0))&"','resource' : '"&JSFormat(formatnumber(oRec("restimer"), 0))&"','startdate' : '"&oRec("jobstartdato")&"','enddate' : '"&oRec("jobslutdato")&"', 'jobresp1' : '"&oRec("jobans1")&"-"&oRec("mnavn")&"-"&oRec("email")&"', 'jobresp2' : '"&oRec("jobans2")&"-"&oRec("mnavn2")&"-"&oRec("email2")&"', 'prioritet' : '"&oRec("risiko")&"', 'datatype' : 'job', 'preconditions_met' : '"& oRec("preconditions_met") &"'" %>}
   
   //AddJobListDB[1] = {'ArrActivities' : ArrActivities, 'ArrMilePael' : ArrMilePael, <%="'customer' : '"&JSFormat(oRec("kkundenavn")) &"' ,'name':'"&JSFormat(oRec("jobnavn"))&"','jnr': '"&oRec("jobnr")&"','jid': '"&oRec("id")&"','cnr': '"&oRec("kkundenr")&"','precalc': '"&JSFormat(timerTildelt)&"','realised': '"&JSFormat(timerforbrugt)&"','resource' : '"&JSFormat(oRec("restimer"))&"','startdate' : '"&oRec("jobstartdato")&"','enddate' : '"&oRec("jobslutdato")&"', 'jobresp1' : '"&oRec("jobans1")&"-"&oRec("mnavn")&"-"&oRec("email")&"', 'jobresp2' : '"&oRec("jobans2")&"-"&oRec("mnavn2")&"-"&oRec("email2")&"', 'prioritet' : '"&oRec("risiko")&"', 'datatype' : 'job'" %>}
   //AddJobListDB[2] = {'ArrActivities' : ArrActivities, 'ArrMilePael' : ArrMilePael, <%="'customer' : '"&JSFormat(oRec("kkundenavn")) &"' ,'name':'"&JSFormat(oRec("jobnavn"))&"','jnr': '"&oRec("jobnr")&"','jid': '"&oRec("id")&"','cnr': '"&oRec("kkundenr")&"','precalc': '"&JSFormat(timerTildelt)&"','realised': '"&JSFormat(timerforbrugt)&"','resource' : '"&JSFormat(oRec("restimer"))&"','startdate' : '"&oRec("jobstartdato")&"','enddate' : '"&oRec("jobslutdato")&"', 'jobresp1' : '"&oRec("jobans1")&"-"&oRec("mnavn")&"-"&oRec("email")&"', 'jobresp2' : '"&oRec("jobans2")&"-"&oRec("mnavn2")&"-"&oRec("email2")&"', 'prioritet' : '"&oRec("risiko")&"', 'datatype' : 'job'" %>}
   
	<%
    c = c + 1
	AddJobListDBnr = AddJobListDBnr + 1
	

                            if lastJob <> oRec("id") then
                            jids = jids & "," & oRec("id")
                            lastJob = oRec("id")

                            jobnrs = jobnrs & ", "& oRec("jobnr")
                            end if

	'response.Flush
	oRec.movenext
	wend
	oRec.close 



  
	


	
	%>
    </script>

   
   
 

   <script type="text/javascript" >
		$(document).ready(function() {
			//Make datepickers
			$.datepicker.setDefaults($.extend({ showOn: 'button', showAnim: 'slideDown', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd/m/yy', changeMonth: true, changeYear: true }));

			$("#FM_ajaxStartDate").datepicker();
			$("#FM_ajaxEndDate").datepicker();
		});
		</script>
	<div id="weekselector">

	</div>


  

  

    	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
    &nbsp;
	


	
	
	
	<%if print <> "j" then

    

	
	            ptop = 40
                pleft = 835
                pwdt = 200

                call eksportogprint(ptop,pleft, pwdt)
                %>
                
                
                    <!--
                     <tr>
                    <td align=center>
                   <a href="job_eksport.asp?jids=<%=jids%>" class=rmenu target="_blank">
                   &nbsp;<img src="../ill/export1.png" border=0 alt="Eksport" /></a>
                </td><td><a href="job_eksport.asp?jids=<%=jids%>" class=rmenu target="_blank">.csv fil eksport</a></td>
                   </tr>
                   
                   -->

                   <form action="job_eksport.asp" method="post" target="_blank">
<tr> <input id="Hidden3" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value="A) Eksportér jobdata >> " style="font-size:9px; width:130px;" />
    <input type="hidden" value="1" name="eksDataStd" /><br />
    <input type="checkbox" value="1" name="xeksDataStd" checked disabled /> Stamdata<br />
    <input type="checkbox" value="1" name="eksDataNrl" /> Nøgletal og Realiseret<br />
    <input type="checkbox" value="1" name="eksDataJsv" /> Jobansvarlige<br />
    <input type="checkbox" value="1" name="eksDataAkt" /> Aktiviteter<br />
    <input type="checkbox" value="1" name="eksDataMile" /> Milepæle
    </td>
</tr>
</form>

                    
                    <tr>
                    <td align=center>
                   <a href="webblik_joblisten21.asp?print=j&media=print" target="_blank">
                   &nbsp;<img src="../ill/printer3.png" border=0 alt="Print version" /></a>
                </td><td><a href="webblik_joblisten21.asp?print=j" target="_blank" class=rmenu>Print version</a></td>
                   </tr>
                  
                    </table>



                    
                     <form action="webblik_joblisten21.asp?func=gemliste" method="post">
                     <input id="Hidden1" name="jobnrs" value="<%=jobnrs%>" type="hidden" />
                     <input id="Hidden2" name="listeid" value="<%=listeid%>" type="hidden" />
                     <input id="Hidden4" name="dtmidt" value="<%=strAar&"/"&strMrd&"/"&strDag%>" type="hidden" />
                    <table width=100% celspacing=0 cellpadding=0>
                        <tr><td>
                        <b>Gem / Opret Gannt liste</b><br />
                        
                        <input type="checkbox" value="1" name="gem_liste" />Opret ny gantt liste som:<br /><input type="text" value="<%=listenavn %>" name="listenavn" style="width:178px; font-size:9px;" /> </td></tr>
                      
                        <tr><td align=right style="padding-right:8px;"><input type="submit" value=" Gem/Opdater >> " style="font-size:9px; width:120px;" /></td></tr>
                          <tr><td style="font-size:9px; color:#999999;"><input type="checkbox" value="1" name="slet_liste" />Slet valgte liste</td></tr>
                    </table>
                    </form>

                    <%if print <> "j" then %>
	            <br /><br /><table cellpadding=0 cellspacing=0 border=0><tr>
                    <td>


                      <%
                nWdt = 120
                nTxt = "Opret nyt job"
                nLnk = "jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>



                    </td></tr></table>
	            
	            <%end if %>
	


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
	
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	<br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
	
	<%
	'Response.write strSQL
	'Response.flush
	%>
<!--#include file="../inc/regular/footer_inc.asp"-->
