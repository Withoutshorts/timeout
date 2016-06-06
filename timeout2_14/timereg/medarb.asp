<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/medarb_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%


media = request("media")


if media <> "eksport" then
'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")


case "FN_getMedlisten"
              
           
             
             
              if len(trim(request.form("cust"))) <> 0 then
              jq_sog = request.form("cust")

               '*** ÆØÅ **'
               'call jq_format(jq_sog)
               'jq_sog = jq_formatTxt
              
              '***** Søger bredt på special karaktere ****
              'jq_sog = replace(jq_sog, "ø", "")
              'jq_sog = replace(jq_sog, "æ", "")
              'jq_sog = replace(jq_sog, "å", "")
              

              'jq_sog = replace(jq_sog, "Ø", "")
              'jq_sog = replace(jq_sog, "Æ", "")
              'jq_sog = replace(jq_sog, "Å", "")

              
              else
              jq_sog = ""
              end if

                


              if jq_sog <> "" then
              sogSQLval = "(mnavn LIKE '%"& jq_sog &"%' OR mnr LIKE '"& jq_sog &"%' OR email LIKE '"& jq_sog &"%' OR init LIKE '"& jq_sog &"%')"
              else
              sogSQLval = "(mid = 0)"
              end if

              'strMedListeTxt = strMedListeTxt & "<table cellspacing='0' cellpadding='0' border='0' width='100%' bgcolor='#EFF3FF'>"
	          'strMedListeTxt = strMedListeTxt & "<tr bgcolor='#5582D2'>"
		      'strMedListeTxt = strMedListeTxt & "<td width='8' rowspan='2' valign='top'>"& sogSQLval &"</td></tr></table>"

              'Response.write strMedListeTxt
              'Response.end
              
              if len(trim(request("jq_visalle"))) <> 0 then
              jq_vispasluk = request("jq_visalle")
              else
              jq_vispasluk = 0
              end if

              if jq_vispasluk <> "1" then
              visAlleSQLval = " AND mansat = 1 "
              else
              visAlleSQLval = " AND mansat <> -1 "
              end if

              lastMid = request("lastmid")
              
              
              '*** Henter medarbejdere ***'
              strMedListeTxt = strMedListeTxt & "<table cellspacing='0' cellpadding='0' border='0' width='100%' bgcolor='#EFF3FF'>"
	          strMedListeTxt = strMedListeTxt & "<tr bgcolor='#5582D2'>"
		      strMedListeTxt = strMedListeTxt & "<td width='8' rowspan='2' valign='top'><img src=""../ill/blank.gif"" width=""8"" height=""31"" border=""0""></td>"
		      strMedListeTxt = strMedListeTxt & "<td colspan=8 valign=""top""><img src=""../ill/blank.gif"" width=""1"" height=""1"" border=""0""></td>"
		      strMedListeTxt = strMedListeTxt & "<td align=right rowspan=""2"" valign=top><img src=""../ill/blank.gif"" width=""8"" height=""31"" border=""0""></td></tr>"
	
	          strMedListeTxt = strMedListeTxt & "<tr bgcolor=""#5582D2""><td height=""30"" width=""200"" class='alt'><b>Navn (Nr.)</b> - Initialer</td>"_
              &"<td class='alt'>Status</td>"_
              &"<td class='alt'>Type</td>"_
	          &"<td class='alt'>Brugergruppe</td>"_
              &"<td class='alt'>Email</td>"_
	          &"<td class='alt'>Sidste login</td>"_
	          &"<td width=""40"" class='alt'>Joblog</td>"_
	          &"<td>&nbsp;</td></tr>"

              

                ''*** Finder medarbejdere i de progrp hvor man er teamleder ***'
	            'call projgrp(-1,level,session("mid"),1)
	    
	             '    medarbgrpIdSQLkri = "AND (mid = "& session("mid")
    
	    
	              '  for p = 0 to prgAntal
	     
	               '  if prjGoptionsId(p) <> 0 then
	                '    call medarbiprojgrp(prjGoptionsId(p), session("mid"))
	                ' end if
	    
	               ' next 
	    
	                ' medarbgrpIdSQLkri = medarbgrpIdSQLkri & ")"
	    
	            'strSQLmids = medarbgrpIdSQLkri '" AND mid = "& usemrn
	        

              
	            strSQL = "SELECT medarbejdere.Mid, Mnavn , Mnr, Mansat, ansatdato, lastlogin, brugergruppe, email, "_
                &" medarbejdertype, type, navn, brugergrupper.id, "_
	            &" medarbejdertyper.id, init, opsagtdato"_
	            &" FROM medarbejdere, brugergrupper, medarbejdertyper "_
	            &" WHERE "& sogSQLval &" "& visAlleSQLval &" AND "_
                &" brugergrupper.id = medarbejdere.brugergruppe AND medarbejdertyper.id = medarbejdere.medarbejdertype"
	
	            '&" l.dato AS lastlogintime "_
	            '&" LEFT JOIN login_historik l ON (l.mid = medarbejdere.mid) "_
	
	            strSQL = strSQL & " GROUP BY medarbejdere.mid"
                strSQL = strSQL & " ORDER BY medarbejdere.mnavn" 
	
	          
	
	
	
	'Response.write strSQL
	'Response.flush

    'Response.end
	
	oRec.open strSQL, oConn, 3
	x = 0
	ak = 0
	while not oRec.EOF 
	

    '*** Rettigheder til at redigere ***'
	if level <= 2 OR level = 6 then
	showallusr = "1" 
	else
		if cint(oRec("Mid")) = cint(session("Mid")) then
		showallusr = "2"
		else
		showallusr = "0"
		end if
	end if
	
	
	
	strMedListeTxt = strMedListeTxt & "<tr><td bgcolor=""#cccccc"" colspan=""9""><img src=""ill/blank.gif"" width=""1"" height=""1"" border=""0""></td></tr>"
	
	
	if cint(lastMid) = oRec("mid") then
	trbgthis = "#ffff99"
	else
	
	
    
        select case right(x,1)
	    case 2,4,6,8,0
	    trbgthis = ""
	    case else
	    trbgthis = "#ffffff"
	    end select
	
	
	end if

    select case oRec("Mansat")
	case "2" 
    msTxt = "Deaktiveret"
	bgSthis = "CRIMSON"
	case "3"
	bgSthis = "#CCCCCC"
    msTxt = "Passiv"
	case else
    bgSthis = "#DCF5BD"
    msTxt = "Aktiv"
	ak = ak + 1
	end select
	
    strMedListeTxt = strMedListeTxt & "<tr bgcolor='"&trbgthis&"'><td>&nbsp;</td><td style=""height:30px;"">"

    select case showallusr 
	case "1", "2"
	strMedListeTxt = strMedListeTxt & "<a href='medarb_red.asp?menu=medarb&func=red&id="& oRec("Mid")&"'>"& oRec("Mnavn") &"</a>"
	case else
    strMedListeTxt = strMedListeTxt & oRec("Mnavn")  
    end select

     if len(oRec("init")) <> 0 then
	    strMedListeTxt = strMedListeTxt & " - [" & oRec("init") & "] "
	 end if

    strMedListeTxt = strMedListeTxt & "<br><span style=""font-size:9px;"">(id: "& oRec("mid")&") "

	strMedListeTxt = strMedListeTxt & "Ansat:"
	if len(oRec("ansatdato")) <> 0 then
	strMedListeTxt = strMedListeTxt & formatdatetime(oRec("ansatdato"),2) 
	end if
	
	
	if len(oRec("opsagtdato")) <> 0 then
	    if formatdatetime(oRec("opsagtdato"),1) <> formatdatetime("1-1-2044",1) then
	    strMedListeTxt = strMedListeTxt & " til " & formatdatetime(oRec("opsagtdato"),2) 
	    end if
	end if
	
    strMedListeTxt = strMedListeTxt & "</span></td><td style='padding:2px 10px 2px 2px; width:80px;'><span style=""background-color:'"&bgSthis&"'; padding:2px; width:65px;"">"& msTxt &"<span></td>"
	strMedListeTxt = strMedListeTxt & "<td>"& oRec("type") &"</td>"
	strMedListeTxt = strMedListeTxt & "<td>"& oRec("navn") &"</td>"
    strMedListeTxt = strMedListeTxt & "<td>"& oRec("email") &"</td>"
	strMedListeTxt = strMedListeTxt & "<td>"& oRec("lastlogin") &"</td>"
	strMedListeTxt = strMedListeTxt & "<td><a href='joblog.asp?menu=timereg&FM_medarb="& oRec("Mid") &"&FM_job=0&selmedarb="& oRec("Mid")&"' class=vmenu>[Se joblog]</a></td>"
	strMedListeTxt = strMedListeTxt & "<td>"
    
    if level <= 2 OR level = 6 then
		
		if lto <> "Demo" AND oRec("Mid") <> 1 then
		strMedListeTxt = strMedListeTxt & "<a href='medarb_red.asp?menu=medarb&func=slet&id="& oRec("Mid") &"'><img src=""../ill/slet_16.gif"" alt=""Slet"" border=""0""></a>"
		end if
		
	else
	strMedListeTxt = strMedListeTxt & "&nbsp;&nbsp;"
	end if
    
    strMedListeTxt = strMedListeTxt & "</td><td valign=""top"" align=""right"">&nbsp;</td></tr>"
	
	x = x + 1
	
	
    'ekspTxt = ekspTxt & ""&x&";xx;cc"
    ekspTxt = ekspTxt & oRec("Mnavn") & ";" & oRec("Mnr") & ";" & oRec("init") & ";" & msTxt & ";" & oRec("type") & ";" & oRec("navn") & ";" & oRec("email") & "; xx99123sy#z"

	oRec.movenext
	wend
	
	strMedListeTxt = strMedListeTxt & "<tr bgcolor=""#5582D2"">"_
	&"<td valign=""top""><img src=""../ill/blank.gif"" width=""8"" height=""10"" border=""0""></td>"_
    &"<td colspan=""8"" valign=""bottom""><img src=""../ill/blank.gif"" width=""1"" height=""1"" border=""0""></td>"_
    &"<td valign=""top"" align=""right""><img src=""../ill/blank.gif"" width=""8"" height=""10"" border=""0""></td></tr>"_	
	&"</table>"

    strMedListeTxt = strMedListeTxt & "<input id=""jq_antal"" type=""hidden"" value="&ak&" />"
    strMedListeTxt = strMedListeTxt & "<input id=""jq_antal_tot"" type=""hidden"" value="&x&" />"

    strMedListeTxt = strMedListeTxt & "<textarea id=""jq_ekspTxt"" style=""display:none; visibility:hidden;"">"&ekspTxt&"</textarea>"
	          
	          '*** ÆØÅ **'
              call jq_format(strMedListeTxt)
              strMedListeTxt = jq_formatTxt
	          
	          Response.Write strMedListeTxt
			

end select
Response.end
end if



end if



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	if len(trim(request("lastmedid"))) <> 0 then
	lastmedid = request("lastmedid")
	else
	lastmedid = 0
	end if
	

    if media <> "eksport" then

	%>
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--include file="../inc/regular/topmenu_inc.asp"-->
	<script src="inc/medarb_jav.js"></script>
	
    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(6)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
    if level = 1 then
	call medarbtopmenu()
    end if
	%>
   
	</div>
        -->

<%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
	
	<%
	


	
	

      if level = 1 then
	%>
	
	
	
	
	
	<%if level <= 2 OR level = 6 then%>
		
		<%
		strSQL = "SELECT licens.key, klienter FROM licens WHERE id = 1"
		oRec.Open strSQL, oConn, 0, 1, 1
		if not oRec.EOF then
		'ikl = left(oRec("key"), 4)
		'akl = right(ikl, 2)
		
		akl = oRec("klienter")
		
		    call antalmedarb_fn(2)

		end if
		oRec.close
		
	end if%>
	
	
	
	
	
	<%if cint(kli) < cint(akl) AND (level = 1) then '* kun admin brugere kan oprette medarb.
	
	
	
	%>   
	        
	
	<%else %>
	<div id="info" style="position:absolute; background-color:#ffffe1; border:1px red dashed; left:400px; top:30px; width:400px; visibility:visible; padding:5px;">
        <img src="../ill/about.gif" /> <b>Info:</b><br />
    Antallet af aktive medarbejdere overstiger det antal medarbejdere i har købt licens til. Kontakt OutZourCE på <a href="mailto:support@outzource.dk" class=vmenu>support@outzource.dk</a> 
	hvis i har spørgsmål vedr. licens og antal medarbejdere.
	</div>
	<%end if%>
	

	
	
	
	<%
	
	if len(trim(request("lastmedid"))) <> 0 then
    lastmid = request("lastmedid")
    else
    lastmid = 0
    end if

    
    
     tTop = 0
	 tLeft = 0
	 tWdth = 800
	       
     
    call tableDiv(tTop,tLeft,tWdth)
	
  
	%>

	
    <form method="post" action="#">
  
    <table cellspacing=2 cellpadding=2 border=0>
    <tr><td colspan="2"> <h4>Medarbejdere <span style="font-size:10px; font-weight:normal;"> - liste</span></h4></td></tr>
    <input type="hidden" id="jq_lastmid" value="<%=lastmid %>" />
    <tr>
        <td valign="top" style="padding-top:30px;"><b>Søg:</b></td>
        <td valign="top" style="padding-top:26px;"><input id="m_sog" type="text" style="width:300px; border:2px #6CAE1C solid;" /><br />
              <span style="color:#999999; font-size:10px;">Navn, Initialer, Nr., Email (% vildcard / vis alle)</span>
        </td>
        <td valign="top" style="padding-top:26px;"><input id="Submit1" type="button" value=" Søg >> " /></td>
        
    
    </tr>
    <tr>
        <td></td>
        <td><input id="jq_vispasogluk" type="checkbox" value="1" /> Vis alle (passive og de-aktiverede)</td>
    </tr>
    
    
    </table>
  

    </div>
    
    <%
	
	 tTop = 20
	 tLeft = 0
	 tWdth = 1020
	 tDsp = ""      
     tVzb = "visible"
     tId = "div_medarb_list"   	
        	
	 call tableDivWid(tTop,tLeft,tWdth,tId, tVzb, tDsp) 
	
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan=8 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
	<td height="30" class='alt'><b>Navn (Nr.)</b> - Initialer</td>
	<td class='alt'>Status</td>
	<td class='alt'>Type</td>
	<td class='alt'>Brugergruppe</td>
    <td class='alt'>Email</td>
	<td class='alt'>Sidste login</td>
	<td width="40" class='alt'>Joblog</td>
	<td>&nbsp;</td>
	</tr>

    </table>
	
    
    <!--- slut table div --->
    </div>
    <br /><br />
	Antal aktive medarbejdere på listen: 
    <input id="antalm" type="text" value="0" style="border:0px; width:50px;" /><br />
    Medarbejdere fundet ialt: <input id="antalmialt" type="text" value="0" style="border:0px; width:50px;" />

    </form>
   

<br><br>
<br>
<a href="Javascript:history.back()"><< Tilbage</a>
<br>
<br>

<%
else

	
	 tTop = 20
	 tLeft = 0
	 tWdth = 800
	       
     call tableDiv(tTop,tLeft,tWdth)

	
  
	%>
    <table><tr><td>Du har ikke adgang til at se denne side.</td></tr></table></div>

    <%

 end if 'level

 else 'media
     

     %>
     <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	 <script src="inc/medarb_jav.js"></script>
	

     <%


'******************* Eksport **************************' 

if media = "eksport" then

    
   
    
    call TimeOutVersion()
    
    
    ekspTxt = request("FM_ekspTxt")
	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
			
				strOskrifter = "Medarbejder; Nr.; Init; Status; Type; Brugergruppe; Email;"
				
				
				
				
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
                
                <!-- 
                onClick="Javascript:window.close()"
                -->



	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_self" >Din CSV. fil er klar >></a>
	            </td></tr>
	            </table>
	            
	          
	            
	            <%
                
                
                Response.end
	            'Response.redirect "../inc/log/data/"& file &""	
				



end if
     
      
     
     end if




if media <> "eksport" AND media <> "print" AND level = 1 then

ptop = 0
pleft = 850
pwdt = 140

call eksportogprint(ptop,pleft,pwdt)
%>

        
        
        
<form action="medarb.asp?media=eksport" method="Post" target="_blank">
<textarea id="ekspTxt" name="FM_ekspTxt" type="text" style="display:none; visibility:hidden;"></textarea>
      <tr>
        <td align=center><img src="../ill/export1.png" border=0>
        </td><td>
        
        <input id="Submit2" type="submit" value=".csv fil eksport >> " style="font-size:9px;" />
        
        </td>
       </tr>
    
        <tr><td colspan="2" style="padding-top:30px;">
          <div style="background-color:forestgreen; padding:5px 5px 5px 5px; width:150px;">
            <a href="medarb_red.asp?menu=medarb&func=opret", " class="alt">Opret ny medarbejder  +</a>
                </div>
            
             </td>
    </tr>
        </form>
   
	
   </table>
</div>
<%end if%>


</div>



<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
