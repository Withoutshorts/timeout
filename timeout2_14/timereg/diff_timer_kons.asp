


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

 <!--#include file="../inc/regular/topmenu_inc.asp"-->


<script>

    $(window).load(function () {

    // run code
    $("#loadbar").hide('fast');
  
});

</script>

<%
'** Variable ***'

   

              if len(trim(request("showresult"))) <> 0 then
              showresult = 1
            else
            showresult = 0
            end if


		    
		    if len(trim(request("FM_frajobnr"))) <> 0 AND cint(showresult) = 1 then
			jobnr_sog = trim(request("FM_frajobnr"))
            show_jobnr_sog = jobnr_sog
            else

            
                if request.cookies("stat")("frajobnr") <> "" then
			    jobnr_sog = request.cookies("stat")("frajobnr")
                show_jobnr_sog = jobnr_sog
                else
                jobnr_sog = "17500 - 19500"
			    jobKri = ""
			    show_jobnr_sog = jobnr_sog 
                end if

          
            end if

      
                
                 
       
			    

       
                if instr(jobnr_sog, ",") <> 0 then '** Komma **'

	                jobKri = " j.jobnr = 0 "
	                jobSogValuse = split(jobnr_sog, ",")
	                for i = 0 TO UBOUND(jobSogValuse)
	                    jobKri = jobKri & " OR j.jobnr = '"& trim(jobSogValuse(i)) &"'"   
	                next
	    
	                jobKri = " AND ("&jobKri&")"

                else

                    if instr(jobnr_sog, "-") <> 0 then '** Interval
	                jobSogValuse = split(jobnr_sog, "-")
	                    
                        call erDetInt(jobSogValuse(0))
                        isInt = isInt
                        call erDetInt(jobSogValuse(1))
                        isInt = isInt  
                                    
                        if isInt = 0 then 'numerisk
                        jobKri = " AND (j.jobnr >= "& trim(jobSogValuse(0)) &" AND j.jobnr <= " & trim(jobSogValuse(1)) & ")"
                        else
                        jobKri = " AND (j.jobnr >= '"& trim(jobSogValuse(0)) &"' AND j.jobnr <= '" & trim(jobSogValuse(1)) & "')"  
                        end if   
                    else
                    
                    if instr(jobnr_sog, ">") <> 0 OR instr(jobnr_sog, "<")  <> 0 then '** Afgræsning
	                
                            
                            if instr(jobnr_sog, ">") <> 0 then
                            jobnr_sog = replace(jobnr_sog, ">", "")
                            
                            call erDetInt(jobnr_sog)
                           
                                    
                            if isInt = 0 then 'numerisk
                            jobKri = " AND (j.jobnr > "& trim(jobnr_sog) &")"  
                            else 
                            jobKri = " AND (j.jobnr > '"& trim(jobnr_sog) &"')"  
                            end if 
                            
                            else
                            
                            jobnr_sog = replace(jobnr_sog, "<", "")

                            call erDetInt(jobnr_sog)
                           
                                    
                            if isInt = 0 then 'numerisk
                            jobKri = " AND (j.jobnr < "& trim(jobnr_sog) &")"   
                            else 
                            jobKri = " AND (j.jobnr < '"& trim(jobnr_sog) &"')" 
                            end if
                        
                            end if
	               
                    else
                    jobKri = "AND (j.jobnr LIKE '"& jobnr_sog &"%' OR j.jobnavn LIKE '"& jobnr_sog &"%') "
                    end if
                    end if

			    end if


           
      
	 
    response.cookies("stat")("frajobnr") = show_jobnr_sog


    

     if len(trim(request("FM_prDato"))) <> 0 then
    prDato = replace(request("FM_prDato"), ".", "-")

    'response.write "prDato:" & prDato
    else
        
        if request.cookies("stat")("prDato") <> "" then
        prDato = request.cookies("stat")("prDato") 
        else
        prDato = day(now)&"-"& month(now)&"-"&year(now)
        end if

    end if

    
    if isDate(prDato) <> true then
    prDato = day(now)&"-"& month(now)&"-"&year(now)
    else
    prDato = prDato
    end if

    response.cookies("stat")("prDato") = prDato


      if len(trim(request("FM_kunde"))) <> 0 OR cint(showresult) = 1 then
    kunde = replace(request("FM_kunde"), ".", "-")

    'response.write "kunde:" & kunde
    else
        
        if request.cookies("stat")("kunde") <> "" then
        kunde = request.cookies("stat")("kunde") 
        else
        kunde = ""
        end if

    end if



    response.cookies("stat")("kunde") = kunde


  

    if len(trim(request("FM_visalle"))) <> 0 then
    visallechk = "CHECKED"
    visalle = 1
    else
    visallechk = ""
    visalle = 0
    end if

%>

	
	
   <!--

	   	    <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	    <%'call tsamainmenu(7)%>
	    </div>
	    <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	    <%
	    'call stattopmenu()
	    %>
	    </div>
       -->


    <%
        
        call menu_2014()
        
        
        if showresult = 1 then%>
      <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:220px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%
	
	exp_loadtid = 0
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	ca.: <b>3-45 sek.</b><br /><br />
    Ved visning af mange job (+300), kan der være lidt ekstra ventetid efter siden er loadet. Først når denne boks forsvinder helt er siden klar.
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>

<%  end if
	
 
	oskrift = "Konsoliderings-oversigt"
	
   
	
        call filterheader_2013(102,20,1000,oskrift)%>
	
	<form action="diff_timer_kons.asp?showresult=1" method="post">
	<table width=100% cellpadding=0 cellspacing=0 border=0>
        <tr><td>
        Kunde: <input type="text" name="FM_kunde" value="<%=kunde %>" style="width:100px;" />
        &nbsp;&nbsp;&nbsp;Vælg jobnr fra <input type="text" name="FM_frajobnr" value="<%=show_jobnr_sog %>" style="width:220px;" /> (maks. 2500)
        &nbsp;&nbsp;&nbsp;Pr. dato: <input type="text" name="FM_prDato" value="<%=prDato %>" style="width:80px;" /> dd-mm-yyyy 
        &nbsp;&nbsp;&nbsp;<input type="submit" value="Søg >>" /></td></tr>
        <tr><td><span style="font-size:9px; color:#999999;">Søg kunde på k% for alle med K. Søg på jobnr, > 17000, eller komma-separeret 17400, 15215, etc.</span></td></tr>
        <tr><td><br /><input type="checkbox" value="1" name="FM_visalle" <%=visallechk %> />Vis også job der stemmer</td></tr>
        </table>
        </form>
         <br /><br /><br />

<%if showresult = 1 then %>


<table width=100% cellpadding=0 cellspacing=0 border=0>
        <tr><td>
    

<%



call akttyper2009(2)

if len(trim(kunde)) <> 0 then 'søg på specifik kunde
strSQL = "SELECT j.id, j.jobnr, j.jobnavn FROM kunder AS k "_
&" LEFT JOIN job AS j ON (j.jobknr = k.kid) WHERE k.kkundenavn LIKE '"& kunde &"%' AND (j.id <> 0 "& jobKri &") AND j.risiko > 0 LIMIT 2500"
else
strSQL = "SELECT j.id, j.jobnr, j.jobnavn FROM job AS j WHERE j.id <> 0 "& jobKri &" AND j.risiko > 0 LIMIT 2500"
end if

'Response.Write strSQL
'Response.flush


%>
<table>
    <tr><td style="width:100px;"><b>Jobnr</b></td><td style="width:250px;"><b>Jobnavn</b></td>
        <td align="right" style="width:150px;"><b>Timer Konsolideret</b><br />Hele måneder, indlæst<br /> pr. 1 i hver måned</td>
        <td align="right" style="width:100px;"><b>Timer Realiseret</b></td>
        <td align="right" style="width:150px;"><b>Konsolider nu</b><br />Job startdato - 12 md.</td></tr>


<%
strJobnr = ""
jobids = "0"
ddato = year(prDato)& "/"& month(prDato) & "/"& day(prDato)
x = 0
oRec.open strSQL, oConn, 3
while not oRec.EOF 
	
    
    	'** timerKons					
	     timerKons = 0							
         strSQLkons = "SELECT sum(timer) AS timerKons FROM timer_konsolideret_tot AS k WHERE k.jobid = " & oRec("id") & " AND dato < '"& ddato &"' GROUP BY jobid"   
        
         'Response.Write "<br>"& strSQLkons
         'Response.flush
                           
         oRec2.open strSQLkons, oConn, 3
         while not oRec2.EOF 

        timerKons = oRec2("timerKons")

            oRec2.movenext
        wend
        oRec2.close


        '** timerT
       
         timerT = 0							
         strSQLtimerT = "SELECT sum(timer) AS timerT FROM timer AS t WHERE t.tjobnr = '"& oRec("jobnr") &"' AND tdato <= '"& year(ddato) &"/"& month(ddato) &"/31""' AND ("& aty_sql_realhours & ") GROUP BY tjobnr"                      
         
        'response.write "<br>"& strSQLtimerT
        'response.flush
        oRec2.open strSQLTimerT, oConn, 3
         while not oRec2.EOF 

        timerT = oRec2("timerT")

            oRec2.movenext
        wend
        oRec2.close


if ((cdbl(timerKons) > cdbl(timerT) + 15) OR (cdbl(timerKons) < cdbl(timerT) - 15)) OR visalle = 1 then
%>
<tr><td><%=oRec("jobnr")%></td><td><%=oRec("jobnavn")%></td><td align="right"><%=formatnumber(timerKons, 2)%></td> <td align="right"><%=formatnumber(timerT, 2)%></td><td align="right" style="white-space:nowrap;">   <a href="mtypgrp_real_konsolider.asp?dothis=1&first=2&jobid=<%=oRec("id") %>" class=vmenu target="_blank">Konsolider job >></a></td></tr>
                             
<%
    if x < 40 then '* .maks 100 job ad gangen
    strJobnr = strJobnr & " OR jobnr = "& oRec("jobnr")
    jobids = jobids & ","& oRec("id")
    end if
    x = x + 1
end if

oRec.movenext
wend
oRec.close



%>
    </table>
            <br /><br /> <br /><br /> <br /><br /> <br /><br /> <br /><br />
Der er <b><%=x %></b> job der ikke stemmer!<br />
Job med +/- mindre end 20 timer bliver ikke vist.<br />
            <br />
            <div style="padding:20px 20px 20px 20px;">
            <b>Konsolider alle:</b> <span style="color:red;">(maks 40, dvs. de 40 første på listen)</span> <br />
            <%=strJobnr %>
            <br /><br />

            <a href="mtypgrp_real_konsolider.asp?dothis=1&first=2&jobid=<%=jobids %>" class=vmenu target="_blank">Konsolider alle ovenstående job >></a>

                </div>
<br><br><br>


&nbsp;
</td></tr>
        </table>

<%end if 'showresult %>

<!--#include file="../inc/regular/footer_inc.asp"-->
