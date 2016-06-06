<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->

 <!--#include file="inc/job_inc.asp"-->



 <% 
     
    '** Jquery section 
    if Request.Form("AjaxUpdateField") = "true" then

  

     Select Case Request.Form("control")
     case "FN_sogjobliste"
              
             
              
              if len(trim(request("sogval"))) <> 0 then
              jq_sog_val = request("sogval")
              else
              jq_sog_val = "XXXXX99999111sdf"
              end if

             level = session("rettigheder")
            

            if level <= 2 then

            strSQLkrijobans = ""

            else

            strSQLkrijobans = " AND (jobans1 = " & session("mid") & " OR jobans2 = " & session("mid") & ")"

            end if
            
        
                     strSQLjob = "SELECT id, jobnavn, jobnr, kkundenavn, kid FROM job "_
                     &"LEFT JOIN kunder as K ON (kid = jobknr) WHERE jobstatus = 1 "& strSQLkrijobans &" AND ((jobnavn LIKE '"& jq_sog_val &"%' OR jobnr LIKE '"& jq_sog_val &"%') "_
                     &" OR (kkundenavn LIKE '"& jq_sog_val &"%' OR kkundenr = '"& jq_sog_val &"'))  ORDER BY kkundenavn, jobnavn"
            
                        lastKid = 0
                        j = 0
                        strjnr = 0
                        strJobOptions = ""

                      oRec.open strSQLjob, oConn, 3
                      while not oRec.EOF 
                        
                        'if j = 0 then
                        
                        'strJobOptions = "<option>Vælg..</option>"
                       
                        'end if

                        if lastKid <> oRec("kid") then
            
                        if j > 0 then
                        strJobOptions = strJobOptions &"<option value='0' DISABLED></option>"
                        end if 

                        strJobOptions = strJobOptions &"<option value='0' DISABLED>"& oRec("kkundenavn") &"</option>"
                        
                        end if
            
                
                        if cint(id) = oRec("id") then
                
                            strSEL = "SELECTED"
                            else
                            strSEL = ""
                
                       end if

                       strJobOptions = strJobOptions &"<option value="& oRec("id") &" "& strSEL &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")</option>"

                       

                       j = j + 1
                       lastKid = oRec("kid")
                      oRec.movenext  
                      wend
                      oRec.close 

                
                    if len(trim(strJobOptions)) = 0 then 
                    strJobOptions = strJobOptions &"<option value='0' DISABLED>Ingen job fundet...</option>"
                    end if
              
            
              '*** ÆØÅ **'
              call jq_format(strJobOptions)
              strJobOptions = jq_formatTxt
	          
	          Response.Write strJobOptions
              


    

    end select
    '** JQery SLUT

     response.end

    end if


     %><script src="inc/jobprintoverblik_jav.js" type="text/javascript"></script><%

     
	func = request("func")
	
    if len(trim(request("id"))) <> 0 then
	id = request("id")
    else
    id = 0
    end if
	

    if len(trim(request("media"))) <> 0 then
    media = request("media")
    else
    media = ""
    end if

    if len(trim(request("visrealtimerdetal"))) <> 0 OR lto = "oko" then
     visrealtimerdetal = 1
     visrealtimerdetalCHK = "CHECKED"
    else
     visrealtimerdetal = 0
     visrealtimerdetalCHK = ""
    end if
    
    showdiv = "stamdata"
	
    thisfile = "jobprintoverblik"
	

        if len(trim(request("sogjob"))) <> 0 then
         sogjobval = request("sogjob")
        else
         sogjobval = ""
        end if

            					
    '** Rediger job ***'

	strSQL = "SELECT j.id, jobnavn, jobnr, kkundenavn, jobknr, "_
	&" jobTpris, jobstatus, jobstartdato, jobslutdato, projektgruppe1, projektgruppe2, "_
	&" projektgruppe3, projektgruppe4, projektgruppe5, j.dato, j.editor, "_
	&" fakturerbart, budgettimer, fastpris, kundeok, j.beskrivelse, "_
	&" ikkeBudgettimer, tilbudsnr, projektgruppe6, projektgruppe7, "_
	&" projektgruppe8, projektgruppe9, projektgruppe10, j.serviceaft, "_
	&" kundekpers, jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, lukafmjob, j.valuta, jobfaktype, rekvnr, "_
	&" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, jo_dbproc, "_
	&" udgifter, risiko, sdskpriogrp, usejoborakt_tp, ski, job_internbesk, abo, ubv, sandsynlighed, "_
    &" diff_timer, diff_sum, jo_udgifter_ulev, jo_udgifter_intern, jo_bruttooms, restestimat, stade_tim_proc, virksomheds_proc, "_
    &" s.navn AS aftalenavn, s.aftalenr, lukkedato"_
    &" FROM job AS j LEFT JOIN kunder AS k ON (k.Kid = jobknr) "_
    &" LEFT JOIN serviceaft AS s ON (s.id = j.serviceaft) WHERE j.id = " & id 
	
	'Response.Write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("jobnavn")
	strjobnr = oRec("jobnr")
	strKnavn = oRec("kkundenavn")
	strKnr = oRec("jobknr")
	strBudget = oRec("jo_bruttooms") 'oRec("jobTpris")
	strStatus = oRec("jobstatus")

    select case strStatus
    case 0
    strStatusNavn = "Lukket"
    case 2
    strStatusNavn = "Passiv"
    case 3
    strStatusNavn = "Tilbud"
    case else
    strStatusNavn = "Aktivt"
    end select
	
    strTdato = oRec("jobstartdato")
	strUdato = oRec("jobslutdato")
    lkdato = oRec("lukkedato")

	strProj_1 = oRec("projektgruppe1")
	strProj_2 = oRec("projektgruppe2")
	strProj_3 = oRec("projektgruppe3")
	strProj_4 = oRec("projektgruppe4")
	strProj_5 = oRec("projektgruppe5")
	strProj_6 = oRec("projektgruppe6")
	strProj_7 = oRec("projektgruppe7")
	strProj_8 = oRec("projektgruppe8")
	strProj_9 = oRec("projektgruppe9")
	strProj_10 = oRec("projektgruppe10")
	strDato = oRec("dato")
	strLastUptDato = oRec("dato")
	strEditor = oRec("editor")
	strFakturerbart = oRec("fakturerbart")
	
	
	strFastpris = oRec("fastpris")
	intkundeok = oRec("kundeok")
	strBesk = oRec("beskrivelse")

    strtilbudsnr = oRec("tilbudsnr")
     
    
    
	intkundekpers = oRec("kundekpers")
    kundekpers = intkundekpers 
	
	jobans1 = oRec("jobans1")
	jobans2 = oRec("jobans2")

    jobans3 = oRec("jobans3")
	jobans4 = oRec("jobans4")

    jobans5 = oRec("jobans5")
	
    jobans_proc_1 = oRec("jobans_proc_1")
    jobans_proc_2 = oRec("jobans_proc_2")
    jobans_proc_3 = oRec("jobans_proc_3")
    jobans_proc_4 = oRec("jobans_proc_4")
    jobans_proc_5 = oRec("jobans_proc_5")

   
	
	'*** Ikke fakturerbare timer bruges ikke mere, men gemmes pga. gamle job ***'
	strBudgettimer = oRec("budgettimer") + oRec("ikkeBudgettimer")
	
	
	
	lukafmjob = oRec("lukafmjob") 
	
    valuta = oRec("valuta")
    jobfaktype = oRec("jobfaktype")
	
	rekvnr = oRec("rekvnr")
	
	jo_gnstpris = oRec("jo_gnstpris")
	jo_gnsfaktor = oRec("jo_gnsfaktor")
	jo_gnsbelob = oRec("jo_gnsbelob")
	jo_bruttofortj = oRec("jo_bruttofortj")
	jo_dbproc = oRec("jo_dbproc")
	
    jo_udgifter_ulev = oRec("jo_udgifter_ulev")
    jo_udgifter_intern = oRec("jo_udgifter_intern")

    jo_bruttooms = oRec("jo_bruttooms")

	udgifter = oRec("udgifter")
	
	intprio = oRec("risiko")
	sdskpriogrp = oRec("sdskpriogrp")
	
	usejoborakt_tp = oRec("usejoborakt_tp")
	
    intServiceaft = oRec("serviceaft")
    aftnavn = oRec("aftalenavn")
    aftnr = oRec("aftalenr")

	ski = oRec("ski")
	abo = oRec("abo")
	ubv = oRec("ubv")
	job_internbesk = oRec("job_internbesk")
	
	intSandsynlighed = oRec("sandsynlighed")

    diff_timer = oRec("diff_timer")
    diff_sum = oRec("diff_sum")
	
    restestimat = oRec("restestimat")
    stade_tim_proc = oRec("stade_tim_proc")

    virksomheds_proc = oRec("virksomheds_proc")

	end if
	oRec.close
	
	
	%>
	

	<!-------------------------------Sideindhold------------------------------------->
	

    <%if level <= 2 OR lto = "oko" then %>
   <div id="sog" style="position:absolute; left:0px; top:10px; visibility:visible; padding:0px 20px 40px 20px; z-index:2000000; width:880px;">
    <form action="jobprintoverblik.asp" method="post">


      
        Søg: <input type="text" id="sogjob" name="sogjob" value="<%=sogjobval %>" style="width:100px;" />
        
        <select name="id" id="jobprint_joblisten" style="width:450px;" onchange="submit();">
           <%

            if len(trim(sogjobval)) <> 0 then
               sogjobvalKri = "AND ((jobnavn LIKE '"& sogjobval &"%' OR jobnr LIKE '"& sogjobval &"%') "_
               &" OR (kkundenavn LIKE '"& sogjobval &"%' OR kkundenr = '"& sogjobval &"'))"
            else
               sogjobvalKri = ""
            end if


            if level <= 2 then

            strSQLkrijobans = ""

            else

            strSQLkrijobans = " AND (jobans1 = " & session("mid") & " OR jobans2 = " & session("mid") & ")"

            end if
            
        
         strSQLjob = "SELECT id, jobnavn, jobnr, kkundenavn, kid FROM job "_
         &"LEFT JOIN kunder as K ON (kid = jobknr) WHERE jobstatus = 1 "& strSQLkrijobans &" "& sogjobvalKri &" ORDER BY kkundenavn, jobnavn"
            
            lastKid = 0
            j = 0
            strjnr = 0

          oRec.open strSQLjob, oConn, 3
          while not oRec.EOF 
            if j = 0 then
            %>
            <option>Vælg job..</option>
            <%
            end if

            if lastKid <> oRec("kid") then
            
            if j > 0 then%>
            <option value="0" DISABLED></option>
            <%end if %>

            <option value="0" DISABLED><%=oRec("kkundenavn")%></option>
            <%
            end if
            
                
            if cint(id) = oRec("id") then
                
                strSEL = "SELECTED"
                else
                strSEL = ""
                
           end if%>

            <option value="<%=oRec("id") %>" <%=strSEL %>><%=oRec("jobnavn") & " ("& oRec("jobnr") &")" %></option>

           <%

           j = j + 1
           lastKid = oRec("kid")
          oRec.movenext  
          wend
          oRec.close %>


            </select>
        <%if lto <> "oko" then %>
        &nbsp;
        <input type="checkbox" name="visrealtimerdetal" value="1" onclick="submit();" <%=visrealtimerdetalCHK %>  /> Vis detal. timeforbrug
        <%end if %>

        &nbsp;&nbsp;&nbsp;<input type="submit" value=" Vis >> " />
    </form>
       </div>
        

  
	<%end if

     
        if id <> 0 then


     %>
<div id="sindhold" style="position:absolute; left:0px; top:0px; visibility:visible; padding:20px 20px 40px 20px;">
<%
     call minioverblik
	
	

   
	
	


    %>
      <br /><br />&nbsp;
    </div>

   <%

   if media = "printjoboverblik" then
   Response.Write("<script language=""JavaScript"">window.print();</script>")
   end if
       
   else 'id <> 0
       
       Response.end
    end if%>
            
               
<!--#include file="../inc/regular/footer_inc.asp"-->
