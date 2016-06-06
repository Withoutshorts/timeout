<%response.buffer = true

 %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="inc/ressource_belaeg_jbpla_inc.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/timbudgetsim_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

<style type="text/css">
    	input { text-align:right; }
        .inputclsLeft { text-align:left; }
</style>


<%

if session("user") = "" then


	errortype = 5
	call showError(errortype)
	else
	


'*****************************************************************************
'***************** VARIABLE **************************************************

if len(trim(request("FM_sog"))) <> 0 then
sogVal = request("FM_sog")
sogValTxt = sogVal
else
sogVal = "WF9)8/NXN76E#"
sogValTxt = ""
end if

 call aktBudgettjkOn_fn()
 fyStMd = month(aktBudgettjkOnRegAarSt)

if len(trim(request("FM_fy"))) <> 0 then
y0 = "01-"&fyStMd&"-"& request("FM_fy")
else
y0 = "01-"&fyStMd&"-"& year(now)
end if

 y1 = dateAdd("yyyy", 1, y0)
 y2 = dateAdd("yyyy", 2, y0) 

 h1aar = datePart("yyyy", dateAdd("yyyy", -1, y0), 2,2) 
 h1md = 7 'juli = Bør sættes udfra: aktBudgettjkOnRegAarSt

 h2aar = datePart("yyyy", y0, 2,2) 
 h2md = 1 'jan


if len(trim(request("FM_visrealprdato"))) <> 0 then
   visrealprdato = request("FM_visrealprdato")

    if isDate(visrealprdato) = false then
    visrealprdato = formatdatetime(now, 2)
    end if

else
   visrealprdato = "1-1-"& h2aar
end if

func = request("func")

'*****************************************************************************


select case func
   case "opd"

    %><div style="position:absolute; left:100px; top:100px; width:200px; border:10px #cccccc solid; background-color:#FFFFFF; padding:20px;"><%
    

    '*** Job ****'
    jobids = split(request("FM_jobid"), ", ")
    jobbudgettimer = split(request("FM_jobbudgettimer"), ", ")
    jobtpris = split(request("FM_jobtpris"), ", ")

    jobBudgetFY0 = split(request("FM_timebudget_FY0"), ", ")
    fctimeprisFY0 = split(request("FM_fctimepris_FY0"), ", ")
    fctimeprish2FY0 = split(request("FM_fctimeprish2_FY0"), ", ")

    jobBudgetFY1 = split(request("FM_timebudget_FY1"), ", ")
    jobBudgetFY2 = split(request("FM_timebudget_FY2"), ", ")

    FY0 = request("FY0")
    FY1 = request("FY1")
    FY2 = request("FY2")

    '*** Medarbejdere på job ***'
    'Response.Write "FM_timebudget_FY0 " & request("FM_timebudget_FY0") & "<br>"
    'Response.Write "FM_fctimepris_FY0 " & request("FM_fctimepris_FY0") & "<br>"
    'Response.Write "FM_fctimeprish2_FY0 " & request("FM_fctimeprish2_FY0") & "<br>"

    'Response.Write "FM_timebudget_FY1 " & request("FM_timebudget_FY1") & "<br>"
    'Response.Write "FM_timebudget_FY2 " & request("FM_timebudget_FY2") & "<br>"

    'response.write "budgettimer: "& request("FM_jobbudgettimer") & "<br>"
    'response.write "jobids: "& request("FM_jobid") & "<br>"
    'response.write "jobtpris: "& request("FM_jobtpris")
    'response.end

  

    for j = 0 TO UBOUND(jobids)

        jobbudgettimer(j) = replace(jobbudgettimer(j), ".", "")
        jobbudgettimer(j) = replace(jobbudgettimer(j), ",", ".")

        if len(trim(jobbudgettimer(j))) = 0 then
        jobbudgettimer(j) = 0
        end if

        jobtpris(j) = replace(jobtpris(j), ".", "")
        jobtpris(j) = replace(jobtpris(j), ",", ".")

        if len(trim(jobtpris(j))) = 0 then
        jobtpris(j) = 0
        end if
        
        '** Hvis job er lbn timer beregn jobTpris / jo_bruttooms
        strSQLupdateJob = "UPDATE job SET budgettimer = " & jobbudgettimer(j) & ", jo_gnstpris = "& jobtpris(j) &" WHERE id = " & jobids(j)
        'response.write strSQLupdateJob
        'response.flush
        oConn.execute(strSQLupdateJob)

        
        for f = 0 to 2
        '** RAMME FY0-FY2 **'

        call opdaterRessouceRamme(f, FY0, FY1, FY2, jobBudgetFY0(j), fctimeprisFY0(j), fctimeprish2FY0(j), jobBudgetFY1(j), jobBudgetFY2(j), jobids(j), 0)

        next


    next

    response.write "Job opdateret...<br><br>"
    response.flush


    medids = split(request("FM_medid"), ", ")
    h1s = split(request("FM_h1"), ", ")
    h2s = split(request("FM_h2"), ", ")

    h1_jobids = split(request("FM_H1_jobid"), ", ")
    h1_aktids = split(request("FM_H1_aktid"), ", ")

    h2_jobids = split(request("FM_H2_jobid"), ", ")
    h2_aktids = split(request("FM_H2_aktid"), ", ")
   

    aar1 = "2014"
    md1 = "7"

    aar2 = "2015"
    md2 = "1"


      for m = 0 TO UBOUND(medids)

                

                if len(trim(h1s(m))) <> 0 then
                h1s(m) = h1s(m)
                else
                h1s(m) = 0
                end if

                if len(trim(h2s(m))) <> 0 then
                h2s(m) = h2s(m)
                else
                h2s(m) = 0
                end if


                h1s(m) = replace(h1s(m), ".", "")
                h1s(m) = replace(h1s(m), ",", ".")

                h2s(m) = replace(h2s(m), ".", "")
                h2s(m) = replace(h2s(m), ",", ".")
        
                fc1Findes = 0

                strSQLdelmedres = "SELECT id FROM ressourcer_md WHERE jobid = " & h1_jobids(m) & " AND aktid = "& h1_aktids(m) &" AND medid = "& medids(m) &" AND aar = "& aar1 &" AND md = "& md1
                oRec3.open strSQLdelmedres, oConn, 3
                if not oRec3.EOF then 

                    
                    if h1s(m) > 0 then
                    strSQLupdmedres = "UPDATE ressourcer_md SET timer = "& h1s(m) &" WHERE id = " & oRec3("id")  
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLupdmedres)
                    end if

                    if h1s(m) = 0 then
                    strSQLdelmedres = "DELETE FROM ressourcer_md WHERE id = " & oRec3("id")  
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLdelmedres)
                    end if


                fc1Findes = 1

                end if
                oRec3.close

            
                '**** Indsætter nye hvsi ikke findes
                if cint(fc1Findes) = 0 then


                    if h1s(m) > 0 then
                    strSQLinsmedres = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES ("& h1_jobids(m) &", "& h1_aktids(m) &", "& medids(m) &", "& aar1 &", "& md1 &", "& h1s(m) &")"  
                    'response.write strSQLinsmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLinsmedres)
                    end if


                end if


                fc2Findes = 0

                strSQLdelmedres = "SELECT id FROM ressourcer_md WHERE jobid = " & h2_jobids(m) & " AND aktid = "& h2_aktids(m) &" AND medid = "& medids(m) &" AND aar = "& aar2 &" AND md = "& md2
                oRec3.open strSQLdelmedres, oConn, 3
                if not oRec3.EOF then 

                    
                    if h2s(m) > 0 then
                    strSQLupdmedres = "UPDATE ressourcer_md SET timer = "& h2s(m) &" WHERE id = " & oRec3("id")  
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLupdmedres)
                    end if


                    if h2s(m) = 0 then
                    strSQLdelmedres = "DELETE FROM ressourcer_md WHERE id = " & oRec3("id")  
                    'response.write strSQLupdmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLdelmedres)
                    end if


                fc2Findes = 1

                end if
                oRec3.close

            
                 '**** Indsætter nye hvsi ikke findes
               

                 if cint(fc2Findes) = 0 then

                     if h2s(m) > 0 then
                    strSQLinsmedres = "INSERT INTO ressourcer_md (jobid, aktid, medid, aar, md, timer) VALUES ("& h2_jobids(m) &", "& h2_aktids(m) &", "& medids(m) &", "& aar2 &", "& md2 &", "& h2s(m) &")"  
                    'response.write strSQLinsmedres & "<br>"
                    'response.flush

                    oConn.execute(strSQLinsmedres)
                    end if

                end if


            next


    response.write "Medarbejdere opdateret...<br><br>"
    response.flush


    '*** Aktiviteter ***'
    aktids = split(request("FM_aktid"), ", ")
    jobaktids = split(request("FM_aktjobid"), ", ")
    aktbudgettimer = split(request("FM_aktbudgettimer"), ", ")
    aktpris = split(request("FM_aktpris"), ", ")

    aktBudgetFY0 = split(request("FM_akttimebudget_FY0"), ", ")
    aktfctimeprisFY0 = split(request("FM_aktfctimepris_FY0"), ", ")
    aktfctimeprish2FY0 = split(request("FM_aktfctimeprish2_FY0"), ", ")

    aktBudgetFY1 = split(request("FM_akttimebudget_FY1"), ", ")
    aktBudgetFY2 = split(request("FM_akttimebudget_FY2"), ", ")

    for j = 0 TO UBOUND(aktids)

        aktbudgettimer(j) = replace(aktbudgettimer(j), ".", "")
        aktbudgettimer(j) = replace(aktbudgettimer(j), ",", ".")

        if len(trim(aktbudgettimer(j))) = 0 then
        aktbudgettimer(j) = 0
        end if


        aktpris(j) = replace(aktpris(j), ".", "")
        aktpris(j) = replace(aktpris(j), ",", ".")

        if len(trim(aktpris(j))) = 0 then
        aktpris(j) = 0
        end if

        '** beregn aktpris pga timepris

        strSQLupdateakt = "UPDATE aktiviteter SET budgettimer = " & aktbudgettimer(j) & ", aktbudget = "& aktpris(j) &" WHERE id = " & aktids(j)
        'response.write strSQLupdateakt
        'response.flush

        oConn.execute(strSQLupdateakt)


        for f = 0 to 2
        '** RAMME FY0-FY2 **'

        call opdaterRessouceRamme(f, FY0, FY1, FY2, aktBudgetFY0(j), aktfctimeprisFY0(j), aktfctimeprish2FY0(j), aktBudgetFY1(j), aktBudgetFY2(j), jobaktids(j), aktids(j))

        next

    next



    response.write "Aktiviteter opdateret...<br><br>"
    response.flush

        
    %>
    Alt er opdateret!<br /><a href="timbudgetsim.asp?FM_fy=<%=FY0 %>&FM_visrealprdato=<%=visrealprdato%>&FM_sog=<%=sogVal %>">Videre..</a>
    </div>
    <%    

   case else    
%>



<script src="inc/timbudgetsim_jav.js"></script>


 <div id="load" style="position:absolute; display:; visibility:visible; top:260px; left:400px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

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

    'if lcase(sogVal) = "all" then

    if len(trim(request("sogmedarb"))) <> 0 then
    medarbinits = request("sogmedarb")
    medarbinitsTXT = medarbinits
    medarbinitSQL = " AND init LIKE '"& medarbinits & "%'"
    else
    medarbinits = ""
    medarbinitSQL = ""
    medarbinitsTXT = ""
     'response.end
    end if

    if lcase(sogVal) <> "all" AND lcase(sogVal) <> "all2"  then
    sogValSQLKri = "AND (kkundenavn LIKE '%"& sogVal &"%' OR jobnavn LIKE '%"& sogVal &"%' OR jobnr LIKE '%"& sogVal &"%') "
    lmt = 20
    pdiv_vzb = "visible"
    pdiv_dsp = ""

    jdiv_vzb = "visible"
    jdiv_dsp = ""

    else
    sogValSQLKri = ""
    lmt = 200


        if lcase(sogVal) <> "all2" then '= all
        jdiv_vzb = "visible"
        jdiv_dsp = ""

            if len(trim(medarbinits)) = 0 then
            pdiv_vzb = "hidden"
            pdiv_dsp = "none"
            else
            pdiv_vzb = "visible"
            pdiv_dsp = ""
            end if
        
        else
        jdiv_vzb = "hidden"
        jdiv_dsp = "none"
        pdiv_vzb = "visible"
        pdiv_dsp = ""
        end if
        
    end if








 p = 14
 m = 120
 mhigh = 0
 phigh = 0
 dim antalm, antalp, h1_medTot, h2_medTot
 redim antalm(m,2), antalp(p), h1_medTot(m), h2_medTot(m) 
 public mhigh, phigh





 %>


 <%call menu_2014() %>
<!-------------------------------Sideindhold------------------------------------->
<div id="sindhold" style="position:absolute; left:90px; top:62px;">


    <h4>Timebudget - Simulering</h4>
<form method="post" action="timbudgetsim.asp">
    Kunde/Job søg: <input class="inputclsLeft" type="text" value="<%=sogValTxt %>" name="FM_sog" placeholder="Søg på kunde ell. jobnavn" style="width:200px;" /> ("all" = viser alle job / ingen forecast) <br />
    Finansår: <select style="width:120px;" name="FM_fy">
        <%for fyC = 0 to 3
            
            if fyC = 0 then
            fyYear = dateAdd("yyyy", -1, now)
            else
            fyYear = dateAdd("yyyy", 1, fyYear)
            end if

            if fyStMd = "7" then
            fyEndYear = year(dateAdd("yyyy", 1, fyYear))
            else
            fyEndYear = year(fyYear)
            end if

            'fyEndDt = dateAdd("d",-1, aktBudgettjkOnRegAarSt)
            'fyEndDt = day(fyEndDt)&". "&left(monthname(month(fyEndDt)), 3)&" "&fyEndYear

            if year(y0) = year(fyYear) then
            ySel = "SELECTED"
            else
            ySel = ""
            end if
             %>
        <option value="<%=year(fyYear) %>" <%=ySel %>>FY <%=year(fyYear) %></option>
        <%next %>

    </select>&nbsp;&nbsp;
    Vis realiseret pr. dato: <input class="inputclsLeft" type="text" name="FM_visrealprdato" value="<%=visrealprdato %>" />
    <input type="submit" value="Søg >>" />  <div style="position:absolute; left:1250px; width:300px; top:60px;">Medarb. init: <input type="text" class="inputclsLeft" value="<%=medarbinitsTXT%>" name="sogmedarb" placeholder="Medarb. init" style="width:100px;" DISABLED />&nbsp; <input type="submit" value="Søg >>" /></div>
</form>


<form method="post" action="timbudgetsim.asp?func=opd&FM_sog=<%=sogVal %>">
    <input type="hidden" name="FY0" value="<%=year(y0)%>" style="width:40px;" />
    <input type="hidden" name="FY1" value="<%=year(y1)%>" style="width:40px;" />
    <input type="hidden" name="FY2" value="<%=year(y2)%>" style="width:40px;" />
    <input type="hidden" name="FM_visrealprdato" value="<%=visrealprdato %>"

<!-- MAIN -->


  <div style="position:absolute; left:0px; top:112px; color:#5582d2; width:200px;" id="a_jobakt">[+] <u>Job og aktiviteter</u></div>
<div id="d_jobakt" style="position:absolute; visibility:<%=jdiv_vzb%>; display:<%=jdiv_dsp%>; left:0px; top:132px; width:1220px; border:1px #999999 solid; background-color:#FFFFFF; padding:10px;">

<h4>Budget</h4>

<!-- Job tabel -->
<table cellpadding="1" cellspacing="0" border="0" width="100%" >

     <thead style="background-color:#5c75AA;">
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>A</td>
        <td>B</td>
        <td>C</td>
        <td>D</td>
        <td>E</td>
      
        <td>F</td>
        <td>G</td>
        <td>H</td>
        <td>I</td>
        <td>I2</td>
        <td>I3</td>
        <td>J</td>
        <td>K</td>
        <td>L</td>
        <td>M</td>
        <td>N</td>
        <td>O</td>
         <td>P</td>

       


    </thead>

    <thead style="background-color:#5c75AA;">
        <td class="alt" valign="bottom">Job</td>
        <td class="alt" valign="bottom">Jobnr.</td>
        <td class="alt" valign="bottom">Aktivitet</td>
        <td class="alt" valign="bottom" style="background-color:#3b5998;">Budgettimer<br />fra job</td>
          <td class="alt" valign="bottom" style="background-color:#3b5998;">Gns.<br />timepris</td>
        <td class="alt" valign="bottom" style="border-bottom:2px red solid;">FY <%=datepart("yyyy", y0, 2,2) %><br />Timebgt.</td>
        <td class="alt" valign="bottom">FY <%=datepart("yyyy", y1, 2,2) %><br />Timebgt.</td>
        <td class="alt" valign="bottom">FY <%=datepart("yyyy", y2, 2,2) %><br />Timebgt.</td>
      
        <td class="alt" valign="bottom">Oprinde<br />lig</td>
        <td valign="bottom" style="background-color:#DCF5BD;">Forecast<br /> timepris<br />H1<br />FY <%=datepart("yyyy", y0, 2,2) %></td>
        <td valign="bottom" style="background-color:#CCCCCC;">Forecast<br /> 1H<br />FY <%=datepart("yyyy", y0, 2,2) %></td>
        <td valign="bottom" style="background-color:#FFFFFF;">Timer<br /> real.<br />FY 1.7 -<br /><b> <%=visrealprdato %></b></td>
        <td valign="bottom" style="background-color:#FFFFFF;">Real. gns. <br />timepris<br /> DKK/t.</td>
        <td valign="bottom" style="background-color:#FFFFFF;">Real. oms.<br />FY 1.7 -<br /><b> <%=visrealprdato %></td>
        <td valign="bottom" style="background-color:#FFFFFF;">Timer til<br /> rådighed<br /> H2<br /><span style="color:#999999;">[C-I]</span></td>
        <td valign="bottom" style="background-color:#DCF5BD;">Forecast<br /> timepris<br />H2<br />FY <%=datepart("yyyy", y0, 2,2) %></td>
        <td valign="bottom" style="background-color:#CCCCCC;">Forecast<br /> 2H<br />FY <%=datepart("yyyy", y0, 2,2) %></td>
        <td valign="bottom" style="background-color:#CCCCCC;">Forecast<br /> H1+H2 <br />FY <%=datepart("yyyy", y0, 2,2) %></td>
        <td valign="bottom" style="background-color:#FFFFFF;">Budget <br />H1 <br />FY <%=datepart("yyyy", y0, 2,2) %><br /><span style="color:#999999;">[G*H]</span></td>
        <td valign="bottom" style="background-color:#FFFFFF;">Budget <br />H2 <br />FY <%=datepart("yyyy", y0, 2,2) %><br /><span style="color:#999999;">[I3 + K*L]</span></td>
        <td valign="bottom" style="background-color:lightpink;">Budget <br />FY <%=datepart("yyyy", y0, 2,2) %><br />

           <span style="color:darkred;"> IF I = 0 : N<br />
            ELSE : O</span>

        </td>

       


    </thead>

<% 

'********************************************************************************* 
'*** Henter job og aktiviteter MAIN  ************************************** *****'
'*********************************************************************************
i = 0

    

lastknavn = ""
lastjobnavn = ""
lastFase = ""
strSQLjob = "SELECT jobnavn, jobnr, j.id AS jid, jobknr, j.budgettimer AS jobbudgettimer, jo_gnstpris, "_
&" kkundenavn, k.kid, a.id AS aid, a.navn AS aktnavn, a.budgettimer AS aktbudgettimer, aktbudget, a.fase FROM job AS j "_
&" LEFT JOIN kunder AS k ON (kid = jobknr) "_
&" LEFT JOIN aktiviteter AS a ON (a.job = j.id) "_
&" WHERE (j.risiko > -1 OR j.risiko = -3) AND j.jobstatus = 1 AND a.aktstatus = 1 "& sogValSQLKri &""_
&" GROUP BY a.id ORDER BY kkundenavn, jobnavn, jobnr, a.fase, a.sortorder, a.navn LIMIT "& lmt

'response.write strSQLjob
'response.Flush
x = 0
oRec.open strSQLjob, oConn, 3
while not oRec.EOF
    

    


    if oRec("jobbudgettimer") <> 0 then
    jobbudgettimer = oRec("jobbudgettimer")
    else
    jobbudgettimer = ""
    end if

    if oRec("jo_gnstpris") <> 0 then
    jobtpris = oRec("jo_gnstpris")
    else
    jobtpris = ""
    end if


    if oRec("aktbudgettimer") <> 0 then
    aktbudgettimer = oRec("aktbudgettimer")
    else
    aktbudgettimer = ""
    end if

    if oRec("aktbudget") <> 0 then
    aktpris = oRec("aktbudget")
    else
    aktpris = ""
    end if


    if lastknavn <> lcase(oRec("kkundenavn")) then%>
    <tr style="background-color:#EFF3FF;"><td colspan="2" style="padding:10px 1px 2px 2px;"><b><%=oRec("kkundenavn") %></b></td>
        <td colspan="2"><!--<input type="submit" value="Opdater >>" />-->&nbsp;</td>
        <td colspan="70">&nbsp;</td></tr>
    <%end if %>
    
            <%if lastjobnavn <> lcase(oRec("jobnavn")) then %>
            <tr>
           
            <input type="hidden" name="FM_jobid" value="<%=oRec("jid") %>" />
            <input type="hidden" id="FM_aktid_<%=i %>" value="0" />
            <input type="hidden" id="FM_jobid_<%=i %>" value="<%=oRec("jid") %>" />
           
            <td style="white-space:nowrap;"><span style="color:#5582d2;" id="sp_<%=oRec("jid") %>" class="sp_jid"><b>[+]</b>&nbsp;&nbsp;</span><%=left(oRec("jobnavn"), 15) %></td><td><%=oRec("jobnr") %></td>
       
             <td>&nbsp;</td>
             <td><input type="text" style="width:60px;" id="budgettimer_jobakt_<%=oRec("jid")%>_0" name="FM_jobbudgettimer" class="jobakt_budgettimer_job" value="<%=jobbudgettimer%>" /></td>
            <td><input type="text" style="width:40px;" id="budgettimep_jobakt_<%=oRec("jid")%>_0" name="FM_jobtpris" class="jobakt_budgettimep_job" value="<%=jobtpris%>" /></td>

                <%call jobaktbudgetfelter(oRec("jobnr"), oRec("jid") , 0, h1aar, h2aar, h1md, h2md) %>

                 <%'call medarbfelter(oRec("jid") , 0, h1aar, h2aar, h1md, h2md) %>
             </tr>
           <% i = i + 1
             x = x + 1
           end if 
               
               
     select case right(x, 1)
    case 0,2,4,6,8
    bgthis = "#FFFFFF"
    case else
    bgthis = "#F3F3F3"
    end select%>

    <%if lastFase <> lcase(oRec("fase")) then %>
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:hidden; display:none;">
        <td colspan="100"><%=replace(oRec("fase"), "_", " ") %></td>
        </tr>
    <%end if %>
        
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:hidden; display:none;">
        <input type="hidden"  name="FM_aktjobid" id="FM_jobid_<%=i %>" value="<%=oRec("jid") %>" />
        <input type="hidden" name="FM_aktid" value="<%=oRec("aid") %>" />
        <input type="hidden" id="FM_aktid_<%=i %>" value="<%=oRec("aid") %>" />
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td style="white-space:nowrap;"><%=left(oRec("aktnavn"), 15) %></td>

        <td><input type="text" name="FM_aktbudgettimer" id="budgettimer_jobakt_<%=oRec("jid")%>_<%=oRec("aid") %>" class="jobakt_budgettimer_job jobakt_budgettimer_<%=oRec("jid")%>" style="width:60px;" value="<%=aktbudgettimer %>" /></td>
           <td><input type="text" style="width:40px;" name="FM_aktpris" id="budgettimep_jobakt_<%=oRec("jid")%>_<%=oRec("aid") %>" class="jobakt_budgettimep_job jobakt_budgettimep_<%=oRec("jid")%>" value="<%=aktpris %>" /></td>


           <%call jobaktbudgetfelter(oRec("jobnr"), oRec("jid"), oRec("aid"), h1aar, h2aar, h1md, h2md) %>


            <%'call medarbfelter(oRec("jid"), oRec("aid"), h1aar, h2aar, h1md, h2md) %>

      
     </tr>
    <%
        if isNull(oRec("fase")) <> true then
        lastFase = lcase(oRec("fase"))
        else
        lastFase = ""
        end if

    i = i + 1
    lastknavn = lcase(oRec("kkundenavn"))
    lastjobnavn = lcase(oRec("jobnavn")) 
    'lastJid = oRec("jid")
    x = x + 1
oRec.movenext    
wend 
oRec.close 


       

%>

    <input type="hidden" value="<%=x-1 %>" id="xhigh" />
    

    <%

    '******************** Medarbejdere LOOP ****************************

        budgetFY0GT = formatnumber(budgetFY0GT, 0)

%>

<tr><td colspan=3>Total:</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td style="background-color:lightpink;"><input type="text" style="width:60px; border:0px; background-color:lightpink;" id="budgetgt" value="<%=budgetFY0GT %>" /></td></tr>

</table>
    <br />
 <input type="submit" value="Opdater >>" />
</div>

   <%if lcase(sogVal) = "all" then 
       
       response.end
       
     end if%>


     <%if lcase(sogVal) = "all2" then
     dvlft = 0
     alft = 150
     else
     dvlft = 1250
     alft = 1250
     end if %>

    <div style="position:absolute; left:<%=alft%>px; top:112px; color:#5582d2; width:200px;" id="a_afdm">[+] <u>Afdelinger / Medarbejdere</u></div>

    <div id="dv_afdm" style="position:absolute; visibility:<%=pdiv_vzb%>; display:<%=pdiv_dsp%>; left:<%=dvlft%>px; top:132px; width:3000px; overflow-x:scroll; border:1px #999999 solid; background-color:#FFFFFF; padding:10px;">
     <h4>Forecast</h4>
<table cellpadding="1" cellspacing="0" border="0">
    
     <thead style="background-color:#5c75AA;"> 


           <%strSQLMedarb = "SELECT mnavn, init, mid, p.navn AS pgrpnavn, p.id AS pid FROM projektgrupper AS p "_
             &"LEFT JOIN progrupperelationer AS pr ON (projektgruppeId = p.id)"_
             &"LEFT JOIN medarbejdere ON (mid = medarbejderId ) WHERE mansat <> 2 "& medarbinitSQL &" AND projektgruppeId <> 10 AND p.id <> 10 ORDER BY pgrpnavn, mnavn LIMIT 80"
            
               'response.write strSQLMedarb
               'response.Flush
              
             lastAfd = ""
             'antalm = 0
               p = 0
               m = 0
             
             oRec5.open strSQLMedarb, oConn, 3
             while not oRec5.EOF
                    
               
                 if lastAfd <> oRec5("pgrpnavn") then%>
          
                 <td valign="bottom" class="alt" style="height:72px;"><span class="sp_p" id="sp_p_<%=p %>"><u><%=left(oRec5("pgrpnavn"), 9) %></u></span></td>
           

            <%

                ' if p = 0 then
                ' pvzb = "visible"
                ' pdsp = ""
                ' else
                 pvzb = "hidden"
                 pdsp = "none"
                ' end if

                antalp(p) = oRec5("pid")
                p = p + 1
                phigh = p
                'm = 0
                end if 


               
                'redim preserve antalm(m,2)
                 'antalm(p,0) = oRec5("pid")
                 antalm(m,0) = oRec5("pid") 
                 antalm(m,1) = oRec5("mid")

                 '** Norm 6 md ***'
                 ntimPer = 0
                 'call normtimerPer(oRec5("mid"), "1-1-"& datepart("yyyy", y0, 2,2), 30, 0)
                 'ntimPer = ntimPer * 6
                 ntimPer = 960
                 antalm(m,2) = ntimPer


                %>
                    
                 <td class="alt afd_p_<%=p-1 %> afd_p" style="white-space:nowrap; visibility:<%=pvzb%>; display:<%=pdsp%>;" valign="bottom"><b><%="[ "& oRec5("init") &" ]" %></b>
                 <br />Tpris</td>
                <td class="alt afd_p_<%=p-1 %> afd_p" valign="bottom" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">H1</td>
                <td class="alt afd_p_<%=p-1 %> afd_p" valign="bottom" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">H2</td>
                <td class="afd_p_<%=p-1 %> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>

             <%

                 'minits(m) = oRec5("init")
                m = m + 1
                mhigh = m
                lastAfd = oRec5("pgrpnavn")

             oRec5.movenext
             wend
             oRec5.close%>


    </thead>

<% 

'***************************************
'*** Henter job og aktiviteter *****'
'***************************************
'i = 0

lastknavn = ""
lastjobnavn = ""
strSQLjob = "SELECT jobnavn, jobnr, j.id AS jid, jobknr, "_
&" kkundenavn, k.kid, a.id AS aid, a.fase FROM job AS j "_
&" LEFT JOIN kunder AS k ON (kid = jobknr) "_
&" LEFT JOIN aktiviteter AS a ON (a.job = j.id) "_
&" WHERE (j.risiko > -1 OR j.risiko = -3) AND j.jobstatus = 1 AND a.aktstatus = 1 "& sogValSQLKri &""_
&" GROUP BY a.id ORDER BY kkundenavn, jobnavn, jobnr, a.fase, a.sortorder, a.navn LIMIT "& lmt

'response.write strSQLjob
'response.Flush
p = 0
x = 0
lastFase = ""
oRec.open strSQLjob, oConn, 3
while not oRec.EOF
    
    if lastknavn <> lcase(oRec("kkundenavn")) then%>
    <tr style="background-color:#EFF3FF;"><td colspan="100" style="padding:10px 1px 2px 2px;">&nbsp;</td></tr>
    <%end if %>
                
            <%'************* Joblinjer ************** %>
            <%if lastjobnavn <> lcase(oRec("jobnavn")) then 
            %><tr><%
            call medarbfelter(oRec("jid") , 0, h1aar, h2aar, h1md, h2md) %>
             </tr>
           <% 'i = i + 1
            end if
               
     select case right(x, 1)
    case 0,2,4,6,8
    bgthis = "#FFFFFF"
    case else
    bgthis = "#F3F3F3"
    end select%>
      <%'************* Aktviitetslinjer ************** %>
    
      <%if lastFase <> lcase(oRec("fase")) then %>
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:hidden; display:none;">
        <td colspan="100">&nbsp;</td>
        </tr>
    <%end if %>
     
    <tr class="tr_<%=oRec("jid") %>" style="background-color:<%=bgthis%>; visibility:hidden; display:none;">
       
     
            <%call medarbfelter(oRec("jid"), oRec("aid"), h1aar, h2aar, h1md, h2md) %>

      
     </tr>
    <% 'i = i + 1
   
    lastknavn = lcase(oRec("kkundenavn"))
    lastjobnavn = lcase(oRec("jobnavn")) 
        
        if isNull(oRec("fase")) <> true then
        lastFase = lcase(oRec("fase"))
        else
        lastFase = ""
        end if
   x = x + 1
oRec.movenext    
wend 
oRec.close 


    '*** totaler ***
     pvzb = "hidden"
     pdsp = "none"
     %>
    <tr><td colspan="100"><br />Total H1 & H2</td></tr>
    <tr>
        <%
      p = 0
      m = 0

     'response.write "mhigh: "& mhigh
     'response.flush

      for p = 0 to phigh - 1
      %>
           
             <td><input type="text" style="width:60px;" disabled /></td>

            <%
        
            for m = 0 TO mhigh - 1 

                'response.write "antalm(m,0) : "& antalm(m,0) & " h1_medTot(antalm(m,1)):"& h1_medTot(antalm(m,1)) &"<br>"
                'response.flush 

                if antalp(p) = antalm(m,0) then

                if antalm(m,2) < h1_medTot(antalm(m,1)) then
                bgThisGT = "red"
                else
                bgThisGT = ""
                end if

                h1h2medTot = (h1_medTot(antalm(m,1))/1 + h2_medTot(antalm(m,1))/1 )%>
                 <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h1_medTot(antalm(m,1)) %>" id="totmedarbh1_<%=m %>" style="width:40px; background-color:<%=bgthisGT%>;" disabled /></td>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=h2_medTot(antalm(m,1)) %>" id="totmedarbh2_<%=m %>" style="width:40px;" disabled /></td>
                <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">= <input type="text" value="<%=h1h2medTot %>" style="width:40px; background-color:#EFF3FF; border:0;" disabled /></td>
                <%
                end if
            next


    next

    %></tr><tr><%


     '*** norm ***
     for p = 0 to phigh - 1
      %>
           
                         <td><input type="text" style="width:60px;" disabled /></td>

                <%
                for m = 0 TO mhigh - 1
                    
                    if antalp(p) = antalm(m,0) then %>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=antalm(m,2) %>" id="totmedarbn1_<%=m %>" style="width:40px;" disabled/></td>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;"><input type="text" value="<%=antalm(m,2) %>" id="totmedarbn2_<%=m %>" style="width:40px;" disabled/></td>
                    <td class="afd_p_<%=p%> afd_p" style="visibility:<%=pvzb%>; display:<%=pdsp%>;">&nbsp;</td>
                    <%
                    end if
                next

     next
    %></tr><%

    '*********************************************************************************************************************************






    %>
    <input type="hidden" value="<%=phigh-1 %>" id="phigh" />
    <input type="hidden" value="<%=mhigh-1 %>" id="mhigh" />

    </table>

    <br />
     <input type="submit" value="Opdater >>" />
    </div>
   






 

   
   
</form>     

    </div><!-- sindhold -->


<%end select 
    
end if 'session%>

<!--#include file="../inc/regular/footer_inc.asp"-->


	
