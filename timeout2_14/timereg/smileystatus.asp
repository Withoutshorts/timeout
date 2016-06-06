<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/smiley_inc.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	thisfile = "smileystatus.asp"
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	'Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
        if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if


    '*** Rettigheder på den der er logget ind **'
	'medarbid = session("mid")
	
	media = request("media")

    if media = "print" then
    print = "j"
    end if

	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then 'ikke længere mulig efer jq vælg alle automatisk
	        
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
	    
        if request.cookies("tsa")("sm_mids") <> "" then
        thisMiduse = request.cookies("tsa")("sm_mids")
        else
	    thisMiduse = session("mid") 
	    end if
        
        intMids = split(thisMiduse, ", ")

	   
	end if

    response.cookies("tsa")("sm_mids") = thisMiduse
	

	if print <> "j" AND media <> "export" AND func <> "slet" AND func <> "sletok" then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
   
	
	<%
    call menu_2014()
    lft = 90
	tp = 72
	else 
	lft = 20
	tp = 20%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if %>



	<div id="sindhold" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; visibility:visible;">
	
	
	<%select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***

    slttxt = "<b>Smiley Status - Slet ugeafslutning?</b><br />"_
	&"Du er ved at slette en <b>ugeafslutning</b>. Er dette korrekt? (du kan altid aflsutte ugen igen)<br><br>"

     strSQL = "SELECT week(u.uge) AS uge, u.afsluttet, week(u.afsluttet) AS afuge, m.mnavn, m.mnr, m.init FROM ugestatus u "_
		 &" LEFT JOIN medarbejdere m ON (m.mid = u.mid) WHERE u.id = " & id
		 
		 'Response.Write strSQL
		 'Response.Flush
		 
		 oRec.open strSQL, oConn, 3
		 if not oRec.EOF then
		 
		 slttxt = slttxt & "<b>Uge: "& oRec("uge") &" - "_
		 & oRec("mnavn") &", ("& oRec("mnr") &") - "& oRec("init") &"</b>" _
		 & "<br>Afsluttet d. "& oRec("afsluttet") & " (uge "& datepart("ww",oRec("afsluttet"),2,2) &")"
		 
		 
		 
		 end if
		 oRec.close

	slturl = "smileystatus.asp?func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,90)

	
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM ugestatus WHERE id = "& id &"")
	
	Response.redirect "smileystatus.asp"
	
	case else
	
	
	%>
	<!-------------------------------Sideindhold------------------------------------->
	
	
	<%
	'if len(request("FM_medarb")) <> 0 then
	'intMid = request("FM_medarb")
	'else
	'intMid = 0
	'end if

    if len(trim(request("FM_yearSel"))) <> 0 then
        yearSel = request("FM_yearSel")
        response.cookies("tsa")("ysel") = yearSel
    else
        if request.cookies("tsa")("ysel") <> "" then
        yearSel = request.cookies("tsa")("ysel")
        else
        yearSel = year(now)
        end if
    end if

        'strDag = "1"
        'strMrd = "1"
        'strAar = year(yearSel)
        'strDag_slut = "31"
        'strMrd_slut = "12"
        'strAar_slut = year(yearSel) 
	

    useYear = yearSel

         pTxt = "Smiley Status (ugeafslutning)"

	if media <> "print" AND media <> "export" then
    
   
    
    call filterheader_2013(32,0,785,pTxt)
    %>
	
    
   <table border=0 cellspacing=2 cellpadding=2 width=100%>
	<form action="smileystatus.asp" method="post" name="periode" id="periode">

        	<%call medkunderjob %>

       

  
	
        <tr>
	
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td valign="top"><br /><b>Periode:</b><br>
			<!--include file="inc/weekselector_s.asp"--> <!-- b -->
                År: <select name="FM_yearSel" onchange="submit();">
                    <%
                        
                        for y = -5 to 10 
                        
                        showyear = yearSel/1 + y
                        if cint(showYear) = cint(yearSel) then
                        ySel = "SELECTED"
                        else
                        ySel = ""
                        end if
                        
                       %>
                    <option value="<%=showYear %>" <%=ySel %>><%=showYear %></option>

                    <%next  %>
                    </select><br />
                <span style="color:#999999; font-size:9px;">Smiley ordningen bliver nulstillet hvert år 1.1</span>
			</td>
	
	
	    
	    
	    <td valign=top style="padding-top:16px; padding-right:20px;" align="right">
	<input type="submit" value="Vis >>" />	</td></tr>
	</form></table>
	
    <!--filter div-->
	</td></tr>
	</table>
	</div>
<br /><br /><br /><br /><br />

<%

ptop = 32
pleft = 825
pwdt = 210

call eksportogprint(ptop,pleft, pwdt) %>
<form method="post" action="smileystatus.asp?media=print&FM_medarb=<%=thisMiduse %>&FM_medarb_hidden=<%=thisMiduse %>&FM_yearsel=<%=yearSel %>" target="_blank">
<table><tr><td> <input type="image" src="../ill/printer3.png"></td><td>
	
	<td><input type="submit" value="Print version >>" style="font-size:9px;" /> </td>
	
   
</td></tr></table>
    </form>

<form method="post" action="smileystatus.asp?media=export&FM_medarb=<%=thisMiduse %>&FM_medarb_hidden=<%=thisMiduse %>&FM_yearsel=<%=yearSel %>" target="_blank">
<table><tr><td> <input type="image" src="../ill/export1.png"></td><td>
	
	<td><input type="submit" value=".csv fil eksport >>" style="font-size:9px;" /> </td>
	
   
</td></tr></table>
    </form>
</div>


   
	<%else 
	
        
         if media <> "export" then%>
	    <!--include file="inc/weekselector_s.asp"--> <!-- b -->
        <h4><%=pTxt %></h4>
	    <b>Periode: <%=yearSel%></b>
        <%end if %>


	<%end if %>
	
	<%
	call smileystatus(thisMiduse, 2, useYear)
	
	
	end select%>
	
	<br /><br />
&nbsp;
	</div>


    <%

      '******************* Eksport **************************' 
                if media = "export" then


    
                    call TimeOutVersion()
    
                 
	                ekspTxt = replace(expTxtsm, "xx99123sy#z", vbcrlf)
	
	                datointerval = request("FM_yearsel")
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\smileystatus.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\afslutperiode_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\afslutperiode_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\afslutperiode_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\afslutperiode_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "afslutperiode_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				               
                                strOskrifter = "Medarbejder; Init; Ansatdato; Uger/Måneder; "

                                    for p = 1 TO ugeMdNrTxtTopKri
                                    strOskrifter = strOskrifter & p &";"
                                    next

                                     
				
				                objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
            
                                <div style="position:absolute; top:60px; left:60px;">				

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
	            
	                            </div>
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if 'media%>


	
	<%if print = "j" then
    Response.Write("<script language=""JavaScript"">window.print();</script>")
    end if%>
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
