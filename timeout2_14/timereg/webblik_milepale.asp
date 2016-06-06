<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->

  <!--#include file="../inc/regular/topmenu_inc.asp"-->

<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	func = request("func")
    thisfile = "webblik_milepale.asp"
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
    media = request("media")
	
	
	
	
    if media <> "print" AND media <> "export" then
        %>

	
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	  
        <!-- 
	    <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	    <%'call tsamainmenu(2)%>
	    </div>
	    <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	    <%
	    'call webbliktopmenu()
	    %>
	    </div>
            -->

     

    <%
          siTop = 102
        siLeft = 90
      
        
        
        call menu_2014()
          
        else 
        
        siTop = 20
        siLeft = 20
        
        %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
     
    
    <%
    
   
    
    
    end if %>

               <div id="sindhold" style="position:absolute; left:<%=siLeft%>px; top:<%=siTop%>px; visibility:visible;">

	
	<%

  

  
	
	'**********************************************************
	'**************** Milepæle *******************************'
	'**********************************************************
	 
	
     if len(trim(request("FM_jobsog"))) <> 0 then
	 
        jobsog = request("FM_jobsog")
	    showjobsog = jobsog
	    jobsogSQL = " AND (j.jobnavn LIKE '"& jobsog &"%' OR j.jobnr = '"& jobsog &"' OR kkundenavn LIKE '"& jobsog &"%' OR kkundenr LIKE '"& jobsog &"')"
	    response.Cookies("webblik")("jobsog") = showjobsog
	 
     else
	        if request.Cookies("webblik")("jobsog") <> "" AND len(request("FM_mitype")) = 0 then
	        jobsog = request.Cookies("webblik")("jobsog")
	        showjobsog = jobsog
	        jobsogSQL = " AND (j.jobnavn LIKE '"& jobsog &"%' OR j.jobnr = '"& jobsog &"' OR kkundenavn LIKE '"& jobsog &"%' OR kkundenr LIKE '"& jobsog &"')"
	        else
	        jobsog = 0
	        showjobsog = ""
	        jobsogSQL = ""
	        response.Cookies("webblik")("jobsog") = showjobsog
	        end if
	 end if
	    
	   
        if len(request("FM_mitype")) <> 0 then
	    typid = request("FM_mitype")
	        
                if typid <> 0 then  
	            typSQL = " m.type = "& typid &" "
	            else
	            typSQL = " m.type <> 0 "
	            end if
	    
        else
	        if request.Cookies("webblik")("mtyp") <> "" then
	        typid = request.Cookies("webblik")("mtyp")
	            if typid <> 0 then  
	            typSQL = " m.type = "& typid &" "
	            else
	            typSQL = " m.type <> 0 "
	            end if
	        else
	        typid = 0
	        typSQL = " m.type <> 0 "
	        end if
	    end if
	    
	    
	    response.Cookies("webblik")("mtyp") = typid
	    response.Cookies("webblik").expires = date + 10
	    
        typSell = ""

		
        
        'if request("FM_ignorer_pg") = "1" then
		'chkThis = "CHECKED"
		'else
		'chkThis = ""
		'end if





	
	if media <> "print" then
	
	 call filterheader_2013(0,0,600,pTxt)


          'oimg = "ikon_milepale_48.png"
	oleft = 0
	otop = 0
	owdt = 600
	oskrift = "Milepæle / Terminer"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)

   %>
	
   <br /><br />
	<form method="post" action="webblik_milepale.asp?FM_usedatokri=1">
                   <table cellpadding="1" cellspacing="2" border="0" width=100%>
	
	<tr>
	    <td valign=top style="padding-top:3px; width:120px;"><b>Søg:</b><br />
	
       <input id="FM_jobsog" name="FM_jobsog" type="text" style="width:300px; border:2px yellowgreen solid;" value="<%=showjobsog %>" /><br />
            <span style="color:#999999; font-size:10px;">Søg på Jobnavn ell. Jobnr, eller kunde.</span>
	    </td>
	<tr>
	<tr>
	    <td style="padding-top:3px;"><b>Type / Status:</b><br />
	    
            <select id="FM_mitype" name="FM_mitype" style="width:200px;" onchange=submit();>
            <option value="0">Alle</option>
         <%strSQLtyper = "SELECT id, navn FROM milepale_typer WHERE id <> 0 ORDER BY id"
	    oRec4.open strSQLtyper, oConn, 3
	     
	    
	    
	    while not oRec4.EOF
	    
	     if cint(typid) = oRec4("id") then
	     typSell = "SELECTED"
	     else
	     typSell = ""
	     end if
	    
	    %>
	    <option value="<%=oRec4("id")%>" <%=typSell%>><%=oRec4("id")%>: <%=oRec4("navn") %></option>
	    <%
	    
	    oRec4.movenext
	    wend
	    oRec4.close
	     %>
	     </select>
	    </td>
	</tr>
	<td valign=top colspan=2 style="padding-top:10px;"><b>Periode:</b></br>
	
		<!--#include file="inc/weekselector_s.asp"-->&nbsp;
		<%
        'if level <= 2 OR level = 6 then %>
		<!--<br><br /><input type="checkbox" name="FM_ignorer_pg" value="1" <%=chkThis%>> Vis job/milepæle uanset tilknytning til dine projektgrupper.-->
		<%'end if%>
		
	
	</td>
	</tr>
	<tr>
	<td colspan=2 align=right style="padding-right:20px;"><br><input type="submit" value="Vis milepæle >> "></td></tr></table>

        </form>
	
	<!-- filter div -->
		</td></tr></table>
	</div>
	
	
	<%
    else
    dontshowDD = 1
	%>
    	<!--#include file="inc/weekselector_s.asp"-->
    <%
    end if 'media

	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
		
	

    if media = "print" then
	Response.write "<br><b>Periode:</b> " & formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)
	end if
	
	
	
    
    tTop = 20
	tLeft = 0
	tWdth = 904
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
	
	
	
	
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	
	<tr><td style="padding-top:0px;">
	
		
		<table cellspacing=0 cellpadding=3 border=0 width=100%>
		<tr bgcolor="#5582d2">
        <td class=alt style="padding:5px 5px 5px 5px;"><b>Milepæl navn</b></td>
	
		<td class=alt style="padding:5px 5px 5px 5px;"><b>Type</b></td>
			<td class=alt style="padding:5px 5px 5px 5px;">Kontakt<br /><b>Job</b></td>
        <td class=alt style="padding:5px 5px 5px 5px;"><b>Slet?</b></td>
		<td class=alt style="padding:5px 5px 5px 5px;"><b>Oprettet af</b></td>
		</tr>
		<%
		'if func = "sletmilepal" then
		''*** Her slettes en milepal ***
		'milepalid = request("milepalid")
		'strSQLmil = "DELETE FROM milepale WHERE id = "& milepalid &""
		'response.write strSQLmil
	'	oConn.execute(strSQLmil)
		'end if
		
		
		
		'datointervalstart = dateadd("d", -5, year(now)&"/"& month(now)&"/"&day(now)) 
		'sqlDatostart = year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
		'datointervalslut = dateadd("d", 65, sqlDatostart)  
		'sqlDatoslut = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
		'"& strPgrpSQLkri &"
		lastDato = ""
        strEksportTxt = ""
		strSQL = "SELECT m.navn, m.milepal_dato, m.jid, m.beskrivelse, j.jobnavn, j.jobnr, m.id AS mileid, "_
		&" mt.id, mt.navn AS typenavn, k.kkundenavn, k.kkundenr, j.id AS jobid, ikon, m.editor AS meditor, m.belob , jobans1, jobans2 "_
		&" FROM milepale m "_
		&" LEFT JOIN job j ON (j.id = m.jid) "_
		&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
		&" LEFT JOIN milepale_typer mt ON (mt.id = m.type) "_
		&" WHERE "& typSQL &" AND (milepal_dato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') "& jobsogSQL &" ORDER BY m.milepal_dato DESC"
		
		'Response.write strSQL
		'Response.flush
		
		oRec.open strSQL, oConn, 3 
		x = 0
		while not oRec.EOF 

        select case right(x, 1)
        case 0,2,4,6,8
        bgt = "#EFF3FF"
        case else
        bgt = "#FFFFFF"
        end select
		
		if len(trim(oRec("jobnavn"))) <> 0 then
		
			if lastDato <> oRec("milepal_dato") then%>
            
            <%if x <> 0 then %>
            <tr bgcolor="#FFFFFF">
			<td valign=bottom colspan="5"><br />&nbsp;</td></tr>
            <%end if %>
			
            <tr bgcolor="#D6Dff5">
			<td valign=bottom colspan="5" style="padding:10px 0px 3px 5px;">
			<span style="font-size:14px;"><b><%=weekdayname(weekday(oRec("milepal_dato"), 2), 2,2) &" "& formatdatetime(oRec("milepal_dato"), 2)%></b></span></td></tr>
			<%
          

		end if%>

        <%if oRec("jobans1") = session("mid") OR oRec("jobans2") = session("mid") OR (oRec("jobans1") = 0 AND oRec("jobans2") = 0) OR level = 1 then
        editok = 1
        else
        editok = 0
        end if %>
			
			<tr bgcolor="<%=bgt %>">

            <td valign=top style="border-top:1px silver solid;" width="250">
            <%if media <> "print" AND editok = 1 then %>
			 <a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("mileid")%>&jid=<%=oRec("jobid")%>&rdir=milep','650','500','250','120');" target="_self"><%=oRec("navn")%>
             <%else %>
             <span style="font-size:12px;"><b><%=oRec("navn")%></b></span>
             <%end if %>
             <%if len(trim(oRec("navn"))) = 0 then %>
             ..
             <%end if %>
             </a>

             <%
               strEksportTxt = strEksportTxt & oRec("milepal_dato") & ";"
             strEksportTxt = strEksportTxt & oRec("navn") & ";"
          %>
			
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<br /><span style="font-size:10px;"><i><%=oRec("beskrivelse")%></i></span>  
			<%
          end if%></td>

			
			
			<td valign=top style="border-top:1px silver solid;">
            <%=oRec("typenavn")%>

            <%strEksportTxt = strEksportTxt & oRec("typenavn") & ";" %>
            
            <%if oRec("belob") <> 0 then %><br />
            <b><%=formatnumber(oRec("belob")) &" "& basisValISO%></b>

            <%
            strEksportTxt = strEksportTxt & formatnumber(oRec("belob")) & ";"
            else
            strEksportTxt = strEksportTxt & "0;"
            end if 
            
            %>

			</td>
			
            <td valign=top style="border-top:1px silver solid;">
			
              <%=left(oRec("kkundenavn"), 30)%>&nbsp;(<%=oRec("kkundenr")%>)<br />

            <%if media <> "print" AND editok = 1 then %>
            <a href="javascript:popUp('jobs.asp?menu=job&func=red&id=<%=oRec("jobid")%>&int=1&rdir=webblik_milepale&nomenu=1','1104','800','20','20');" target="_self" class=vmenu><%=left(oRec("jobnavn"), 30)%> (<%=oRec("jobnr")%>)</a>
			<%else %>
            <b><%=oRec("jobnavn")%></b> (<%=oRec("jobnr")%>)
            <%end if %>

            <%strEksportTxt = strEksportTxt & oRec("jobnavn") &";"& oRec("jobnr") &";"& oRec("kkundenavn") &";"&oRec("kkundenr") &";"
            
            if len(trim(oRec("beskrivelse"))) <> 0 then
              call  htmlparseCSV(oRec("beskrivelse"))
            strJobBesk = htmlparseCSVtxt 
            strEksportTxt = strEksportTxt & Chr(34) & strJobBesk & Chr(34) & ";"
            else
            strEksportTxt = strEksportTxt & ";"
            end if
            
            %>

            


            </td>
			

            <td valign=top style="border-top:1px silver solid;">
            <%if media <> "print" AND editok = 1 then %>
            <a href="javascript:popUp('milepale.asp?menu=job&id=<%=oRec("mileid")%>&jid=<%=jid%>&func=slet&rdir=milep','650','500','250','120');" target="_self" class="red">[x]</a>
            <%else %>
            &nbsp;
            <%end if %>
            </td>
			
			<td valign=top style="border-top:1px silver solid;">
			<span style="font-size:10px; color:#999999;"><i><%=oRec("meditor") %></i></span>
			</td>
			
			
			
			</tr>
			
			<%
            strEksportTxt = strEksportTxt & "xx99123sy#z"
			lastDato = oRec("milepal_dato")
			x = x + 1
		
		end if
		
		oRec.movenext
		wend
		oRec.close 
		%>
		</table>
		</td></tr></table>
	    <br /><span style="color:#999999;">NB: Milepæle for Job start- og slut -datoer bliver ikke vist på listen.</span>
	   
	
	
	
	<!-- table div slut -->
		</div>



        <%if media = "export" then 

    

	ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	call TimeOutVersion()
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")

                file = "milepaleexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\webblik_milepale.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\"&file, True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\"&file, 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\"&file, True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\"&file, 8)
				end if
				
				
				
				
				
				
				'**** Eksport fil, kolonne overskrifter ***
				strOskrifter = "Dato;Milepæl;Type;Beløb;Job;Jobnr;Kunde;Kundenr;Beskrivelse;"

				'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				objF.WriteLine(strOskrifter & vbcrlf)
				objF.WriteLine(ekspTxt)
				

               
                Response.redirect "../inc/log/data/"& file &""	

               Response.Write("<script language=""JavaScript"">window.close();</script>")
              
				
	Response.end
	

end if %>



            <%  

if media <> "print" then



ptop = 0
pleft = 640
pwdt = 220

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="webblik_milepale.asp?media=export" method="post" target="_blank">
<tr> 
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value="A) Eksporter milepæle >> " style="font-size:9px; width:130px;" /><br />
    Eksporter viste milepæle som .csv fil</td>
</tr>
</form>



<form action="webblik_milepale.asp?media=print" method="post" target="_blank">
<tr>

    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td><td class=lille><input id="Submit6" type="submit" value="B) Print milepæle >>" style="font-size:9px; width:130px;" /><br />
    Print version</td>
</tr>
</form>


<tr><td colspan="2"><br /><br />


    <%
                nWdt = 120
                nTxt = "Opret ny milepæl"
                nLnk = "milepale.asp?menu=job&func=opr&jid=0&rdir=milep"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>
    </td></tr>

   
	
   </table>
</div>
	<%
    
    else
    
     if media = "print" then
        Response.Write("<script language=""JavaScript"">window.print();</script>")
    end if

	end if%>


	 <br /><br /><br /><br />&nbsp;
	
	
	
	
	
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->