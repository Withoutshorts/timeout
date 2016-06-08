




<%
    '*** Salg og Værdi SQL kald ***'
  public strSQLjob
  sub salgogvaerdiSQL


  'Response.write "her"
  'Response.flush

    strSQLjob = "SELECT j_td1.jobnavn, j_td1.id AS jid, j_td1.jobnr, j_td1.serviceaft, "_
    &" "& formelSQL &", j_td1.jobans_proc_1 AS td1_proc1, j_td1.jobans_proc_2 AS td1_proc2, j_td1.jobans_proc_3 AS td1_proc3, j_td1.jobans_proc_4 AS td1_proc4, "_
    &" j_td1.jobans_proc_5 AS td1_proc5, "_
    &" j_td1.virksomheds_proc AS td1_virksomhedsproc, k.kkundenavn "_
    &" FROM job AS j_td1 "_
    &""& fakSQL &""_
    &" LEFT JOIN valutaer AS v ON (v.id = "& joinValutaFlt &") "_
    &" LEFT JOIN kunder AS k ON (k.kid = jobknr) "_
    &" WHERE ((j_td1.jobans1 = "& oRec2("mid") &" OR j_td1.jobans2 = "& oRec2("mid") &" OR j_td1.jobans3 = "& oRec2("mid") &" OR j_td1.jobans4 = "& oRec2("mid") &" OR j_td1.jobans5 = "& oRec2("mid") &") AND "_
    &""& jobStKriSQL &") "& whEkstra &" "_
    &" GROUP BY "& sqlGrpBy &" ORDER BY j_td1.id "

    end sub

       
   

    '*** DB proent *** 
    public dbProc
    function fn_dbproc(omsbel,dbbel)

    '* dbbel = Samlet oms - udgifter 

    '** dbbel vil atid være >= 0 da det er en delmængde af det samlede beløb
    if omsbel <> 0 then

            dbProc = ((dbbel * 100)/omsbel)
           
    else
    dbProc = 0
    end if

    end function


     


       public restimerThis, li
       function resTimer(jobid, aktid, medid, md, aar, md_split_cspan, saldo)


       'Response.Write "medid" & medid & "<br>"

       if md_split_cspan = 1 OR medid = 0 then

       
               mdSt = month(datoStart)
               aarSt = year(datoStart)
               mdSl = month(datoSlut)
               aarSL = year(datoSlut)

               if year(datoStart) = year(datoSlut)  then
	           orandval = " AND "
	           else
	           orandval = " OR "
	           end if

               if medid = 0 then 'jobtotal
               medidKri = replace(medarbSQlKri, "m.mid", "rmd.medid")
               grpByKri = "rmd.jobid"
               else
               medidKri = "rmd.medid = "& jobmedtimer(x,4) &""
               grpByKri = "rmd.jobid, rmd.medid"
               end if

               select case saldo 
               case 2 ' Saldo uanset periode
               datoKri = "(rmd.md >= 1 AND rmd.aar >= 2001)"
               case 0 '*** Saldo periode
               datoKri = "((rmd.md >= "& mdSt &" AND rmd.aar = "& aarSt & ") "& orandval &" (rmd.md <= "& mdSl &" AND rmd.aar = "& aarSl &"))"
               case 1 '** Prev saldo
               datoKri = "(rmd.md < "& mdSt &" AND rmd.aar <= "& aarSt & ")"
               end select
       


       else
       md = md
       aar = aar
       datoKri = "rmd.md = "& md &" AND rmd.aar = "& aar 

       medidKri = "rmd.medid = "& jobmedtimer(x, 4) &""
       grpByKri = "rmd.jobid, rmd.medid"
       end if


       if aktid <> 0 then 'aktivtet valgt
       jobaktIDKri = "rmd.aktid = "& aktid
       else
       jobaktIDKri = "rmd.jobid = "& jobid
       end if

       restimerThis = 0

        strSQLrestimer = "SELECT sum(timer) AS restimer FROM ressourcer_md AS rmd WHERE "& jobaktIDKri &""_
        &" AND ("& medidKri &") AND ("& datoKri &") GROUP BY "& grpByKri &""

        'if medid <> 0 then
        'Response.Write li &" (lastMD "& LastMd &"): "& strSQLrestimer & "<br>"
        'Response.flush
        'end if
			                        
        oRec4.open strSQLrestimer, oConn, 3
        if not oRec4.EOF then
        restimerThis = oRec4("restimer")
        end if
        oRec4.close

      if resTimerThis <> 0 then
      restimerThis = formatnumber(restimerThis,0)
      else
      restimerThis = ""
      end if

       end function
       

        
       


function sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	<div id="sideoverskrift" style="position:relative; left:<%=oleft%>px; top:<%=otop%>px; width:<%=owdt%>px; visibility:visible; border:0px #000000 solid; z-index:1200;">
	<!-- sideoverskrift -->
	<table cellspacing="0" cellpadding="5" border="0" width=100%>
	<tr>
         <%if media <> "print" then %>
	    <td style="width:48px;"><img src="../ill/<%=oimg%>" alt="" border="0"></td>
        <td align=left style="padding-top:30px; padding-left:20px;">
        <%else %>
        <td align=left style="padding-top:30px; padding-left:0px;">
        <%end if %>
	    <h3><%=oskrift %>


            <%select case lto
             case "mmmi", "intranet - local"
             case else %>
        <%if media = "print" then %>
        <span style="color:#999999; font-size:10px; line-height:15px; font-style:normal; font-weight:normal;"><br /><b>TimeOut</b> <%=formatdatetime(now,1) & " "& formatdatetime(now,3) %>&nbsp;
        Zoom: [CTRL] + [+/-]</span>
        <%end if 
            
            end select
            %>
        </h3>


        </td>
	</tr>
	</table>
	</div>	
	<%
	end function



    function sideoverskrift_2014(oleft, otop, owdt, oskrift)
	%>
	<div id="Div2" style="position:relative; left:<%=oleft%>px; top:<%=otop%>px; width:<%=owdt%>px; visibility:visible; border:0px #000000 solid; z-index:100;">
	<!-- sideoverskrift -->
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<tr>
       
	
        <td align=left>
       
	    <h4><%=oskrift %>


            <%select case lto
             case "mmmi", "intranet - local"
             case else %>
        <%if media = "print" then %>
        <span style="color:#999999; font-size:10px; line-height:15px; font-style:normal; font-weight:normal;"><br /><b>TimeOut</b> <%=formatdatetime(now,1) & " "& formatdatetime(now,3) %>&nbsp;
        Zoom: [CTRL] + [+/-]</span>
        <%end if 
            
            end select
            %>
        </h4>


        </td>
	</tr>
	</table>
	</div>	
	<%
	end function


    
    
    public chkMedarblinier, chklog, chktlffax, chkemail, chkdeak, chklukjob, ign_akttype_inst
	function setFakPreInd()
	
	
	if len(request("FM_chkmed")) <> 0 then
	    if request("FM_chkmed") = 1 then
	    chkMedarblinier = 1
	    else
	    chkMedarblinier = 0
	    end if
	else
	    if len(trim(request.cookies("erp")("tvmedarb"))) <> 0 then
	    chkMedarblinier = request.cookies("erp")("tvmedarb")
	    else
        chkMedarblinier = 0
	    end if
	end if
	


	Response.Cookies("erp")("tvmedarb") = chkMedarblinier
	
	
	if len(request("FM_chklog")) <> 0 then
	    if request("FM_chklog") = 1 then
	    chklog = 1
	    else
	    chklog = 0
	    end if
	else
	    chklog = request.cookies("erp")("tvlogs")
	end if
	
	if len(request("FM_chktlffax")) <> 0 then
	    if request("FM_chktlffax") = 1 then
	    chktlffax = 1
	    else
	    chktlffax = 0
	    end if
	else
	    chktlffax = request.cookies("erp")("tvtlffax")
	end if
	
	if len(request("FM_chkemail")) <> 0 then
	    if request("FM_chkemail") = 1 then
	    chkemail = 1
	    else
	    chkemail = 0
	    end if
	else
	    chkemail = request.cookies("erp")("tvemail")
	end if
	
	'if len(request("FM_chkdeak")) <> 0 then
	'    if cint(request("FM_chkdeak")) = 1 then
	'    chkdeak = 1
	'    else
	    chkdeak = 0
	'    end if
	'else
	'    if request.cookies("erp")("deak") <> "" then
	'    chkdeak = request.cookies("erp")("deak")
	'    else
	'    chkdeak = 0
	'    end if
	'end if
	
	if len(request("FM_chklukjob")) <> 0 then
	    if cint(request("FM_chklukjob")) = 1 then
	    chklukjob = 1
	    else
	    chklukjob = 0
	    end if
	else
	    if request.cookies("erp")("lukjob") <> "" then
	    chklukjob = request.cookies("erp")("lukjob")
	    else
	    chklukjob = 0
	    end if
	end if

    '** ignorer akttype inst. **'
    if len(request("FM_chkmed")) <> 0 then '** Er form submitted **'
        if len(trim(request("ign_akttype_inst"))) <> 0 then
        ign_akttype_inst = request("ign_akttype_inst")
        else
        ign_akttype_inst = 0
        end if
    else
        if request.cookies("erp")("ignakttype") <> "" then
	    ign_akttype_inst = request.cookies("erp")("ignakttype")
	    else
	    ign_akttype_inst = 0
	    end if
        
    end if
	
	Response.cookies("erp")("lukjob") = chklukjob 
	Response.cookies("erp")("deak") = chkdeak 
	Response.Cookies("erp")("tvemail") = chkemail
	Response.Cookies("erp")("tvtlffax") = chktlffax
	Response.Cookies("erp")("tvlogs") = chklog
	Response.Cookies("erp")("ignakttype") = ign_akttype_inst

	end function
    
    

    public  betCHK9, betCHK8, betCHK7,  betCHK6,  betCHK5,  betCHK4,  betCHK3,  betCHK2,  betCHK1, betCHKu1
    public betCHK10, betCHK11, betCHK12, betCHK13, betCHK14, betCHK15, betCHK16, betCHK20, betCHK21, betCHK22, betbetint_txt 
	function betalingsbetDage(betbetint,disa,lang,nameid)
	
	
	    
        if cint(lang) = 1 then 'UK

        betTxt9 = "60 days net"
	    betTxt8 = "Current month plus 30 days net"
	    betTxt7 = "Current month plus 45 days net"
	    betTxt6 = "45 days net"
	    betTxt5 = "21 days net"
	    betTxt4 = "Current month plus 15 days net"
	    betTxt2 = "14 days net"
	    betTxt3 = "30 days net"
	    betTxt1 = "8 days net"
	    betTxtu1 = "Choose at invoicing"
        betTxt0 = "Choose.."
        betTxt20 = "0 days net"
        betTxt21 = "0 days NetCash"
        betTxt10 = "Current month plus 63 days net"
        betTxt11 = "Current month plus 60 days net"
        betTxt12 = "Current month plus 90 days net"
        betTxt13 = "Current month plus 120 days net"
        betTxt14 = "90 days net"
        betTxt15 = "120 days net"
        betTxt16 = "Current month plus 62 days net"
        betTxt22 = "Current month plus 20 days net"

        else


        betTxt9 = "60 dage"
	    betTxt8 = "Lbn. månd + 30 dage"
	    betTxt7 = "Lbn. månd + 45 dage"
	    betTxt6 = "45 dage"
	    betTxt5 = "21 dage"
	    betTxt4 = "Lbn. månd + 15 dage"
	    betTxt2 = "14 dage"
	    betTxt3 = "30 dage"
	    betTxt1 = "8 dage"
	    betTxtu1 = "Angiver selv"
        betTxt0 = "Vælg.. (ikke angivet)"
        betTxt20 = "0 dage"
        betTxt21 = "0 dage NetCash"
        betTxt10 = "Lbn. månd + 63 dage"
        betTxt11 = "Lbn. månd + 60 dage"
        betTxt12 = "Lbn. månd + 90 dage"
        betTxt13 = "Lbn. månd + 120 dage"
        betTxt14 = "90 dage"
        betTxt15 = "120 dage"
        betTxt16 = "Lbn. månd + 62 dage"
        betTxt22 = "Lbn. månd + 20 dage"


        end if

        betCHK9 = ""
	    betCHK8 = ""
	    betCHK7 = ""
	    betCHK6 = ""
	    betCHK5 = ""
	    betCHK4 = ""
	    betCHK2 = ""
	    betCHK3 = ""
	    betCHK1 = ""
	    betCHKu1 = ""
        betCHK0 = ""
        betCHK20 = ""
        betCHK21 = ""
        betCHK10 = ""
        betCHK11 = ""
        betCHK12 = ""
        betCHK13 = ""
        betCHK14 = ""
        betCHK15 = ""
        betCHK16 = ""
        betCHK22 = ""
        betbetint_txt = ""
	
	select case betbetint
	    case 20
        betCHK20 = "SELECTED"
        betbetint_txt = betTxt20
        case 21
        betCHK21 = "SELECTED"
        betbetint_txt = betTxt21
        case 22
        betCHK22 = "SELECTED"
        betbetint_txt = betTxt22
        case 1
	    betCHK1 = "SELECTED"
	    betbetint_txt = betTxt1
        case 2
	    betCHK2 = "SELECTED"
        betbetint_txt = betTxt2
	    case 3
	    betCHK3 = "SELECTED"
        betbetint_txt = betTxt3
	    case 4
	    betCHK4 = "SELECTED"
        betbetint_txt = betTxt4
	    case 5
	    betCHK5 = "SELECTED"
        betbetint_txt = betTxt5
	    case 6
	    betCHK6 = "SELECTED"
        betbetint_txt = betTxt6
	    case 7
	    betCHK7 = "SELECTED"
        betbetint_txt = betTxt7
	    case "-1"
	    betCHKu1 = "SELECTED"
        betbetint_txt = betTxtu1
	    case 8
	    betCHK8 = "SELECTED"
        betbetint_txt = betTxt8
	    case 9
	    betCHK9 = "SELECTED"
        betbetint_txt = betTxt9
	    case 10
	    betCHK10 = "SELECTED"
        betbetint_txt = betTxt10
        case 11
	    betCHK11 = "SELECTED"
        betbetint_txt = betTxt11
        case 12
	    betCHK12 = "SELECTED"
        betbetint_txt = betTxt12
        case 13
	    betCHK13 = "SELECTED"
        betbetint_txt = betTxt13
        case 14
	    betCHK14 = "SELECTED"
        betbetint_txt = betTxt14
        case 15
	    betCHK15 = "SELECTED"
        betbetint_txt = betTxt15
        case 16
	    betCHK16 = "SELECTED"
        betbetint_txt = betTxt16
	    end select 
        
        if cint(disa) = 1 then
        selFFdatoDis = "DISABLED"
        else
        selFFdatoDis = ""
        end if
	  
	    %>
        <select id="<%=nameid %>" name="<%=nameid %>" class="form-control input-small" <%=selFFdatoDis %>>
		

                <option value="0"><%=betTxt0%></option>
            <option value="-1" <%=betCHKu1 %>><%=betTxtu1%></option>
                <option value="20" <%=betCHK20 %>><%=betTxt20%></option>
                <option value="21" <%=betCHK21 %>><%=betTxt21%></option>
                <option value="1" <%=betCHK1 %>><%=betTxt1%></option>
                <option value="2" <%=betCHK2 %>><%=betTxt2%></option>
                <option value="5" <%=betCHK5 %>><%=betTxt5%></option>
                <option value="3" <%=betCHK3 %>><%=betTxt3%></option>
                 <option value="6" <%=betCHK6 %>><%=betTxt6%></option>
                 <option value="9" <%=betCHK9 %>><%=betTxt9%></option>
                 <option value="14" <%=betCHK14 %>><%=betTxt14%></option>
                <option value="15" <%=betCHK15 %>><%=betTxt15%></option>

                <option value="4" <%=betCHK4 %>><%=betTxt4%></option>
                <option value="22" <%=betCHK22 %>><%=betTxt22%></option>
                <option value="8" <%=betCHK8 %>><%=betTxt8%></option>
                <option value="7" <%=betCHK7 %>><%=betTxt7%></option>
                
                <option value="11" <%=betCHK11 %>><%=betTxt11%></option>
                <option value="16" <%=betCHK16 %>><%=betTxt16%></option>
                <option value="10" <%=betCHK10 %>><%=betTxt10%></option>
                
                <option value="12" <%=betCHK12 %>><%=betTxt12%></option>
                <option value="13" <%=betCHK13 %>><%=betTxt13%></option>
                
            </select>
	
	<%
	end function

public antalFak
function antalFakturaerKid(kid)
'*** kunde må kun slettes hvis der ikke findes fakturaer **'
			antalfak = 0
			strSQLfak = "SELECT COUNT(fid) AS antalfak FROM fakturaer WHERE fakadr = " & kid & " GROUP BY fakadr"
			oRec4.open strSQLfak, oConn, 3
			if not oRec4.EOF then
			
			antalfak = oRec4("antalfak")
			
			end if
			oRec4.close 
end function


	
	

	
	
    function sltque(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="slet" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffFF; width:300px; visibility:visible; border:10px #fb1c3e solid; padding:20px;">
	<table cellspacing="0" cellpadding="4" border="0">
	<tr>
	    <td ><h4>Slet?</h4> </td>
	    <td align=right>&nbsp;</td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %> style="color:red;">Ja - slet denne</a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %> style="color:red;">Ja - slet denne</a>
		</td>
		<td align=right>
		<a href="Javascript:history.back()">Nej, tilbage</a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	 function sltque_Small(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="Div4" style="position:absolute; width:253px; left:<%=lft%>px; top:<%=tp%>px; padding:10px; background-color:#FFFFFF; visibility:visible; border:1px #8cAAe6 solid;">
	<table cellspacing="0" cellpadding="2" border="0" width=100%>
	<tr>
	    <td colspan="2"><h4>Slet?</h4></td>
	   
    </tr>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	 
         <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - nulstil" border="0"></a>
		</td>
	    </tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 style="padding:5px;"><%=slttxt %><br /><br />&nbsp;
	    </td>
	   </tr>
	  <tr><td style="padding-left:5px;">
		<a href="Javascript:history.back()" class=rmenu><< Nej, tilbage</a><br />&nbsp;
		</td>
	    <td align=right style="padding-right:25px;">
		<a href=<%=slturl %> class=red>Ja, slet denne >></a><br />&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function sltquePopup(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="Div1 class="jcorner" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffff; visibility:visible; border:10px red solid; padding:20px;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#ffffff">
	<tr>
	    <td bgcolor="#ffffff" colspan="2"><h4>Slet?</h4> </td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffff"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %> style="color:red;">Ja - slet</a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffff"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %> style="color:red;">Ja - slet</a>
		</td>
		<td align=right>
		<a href="Javascript:window.close()">Nej tilbage</a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function eksportogprint(ptop,pleft,pwdt)
        if pwdt < 180 then
        pwdt = 180
        end if
	%>
	<div id=eksport style="position:absolute; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; padding:10px 10px 10px 10px; z-index:1000;">
  
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr>
    <td height=30 bgcolor="#ffffff" style="border-bottom:0px #cccccc solid;" colspan=2><h4>Print & Funktioner</h4></td>
    </tr>
	<%
	end function
	
    
     function eksportogprint_float(ptop,pleft,pwdt)
        if pwdt < 180 then
        pwdt = 180
        end if
	%>
	<div id=eksport style="background-color:#ffffff; width:<%=pwdt%>px; z-index:2000000; float:right; padding:10px 10px 10px 10px;">
  
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr>
    <td height=30 bgcolor="#ffffff" style="border-bottom:0px #cccccc solid;" colspan=2><h4>Print & Funktioner</h4></td>
    </tr>
	<%
	end function

	function eksportogprint09(ptop,pleft,pwdt)
	%>
	<div id=eksport style="position:relative; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; padding:3px 3px 3px 3px;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr>
    <td bgcolor="#FFFFFF" style="border-bottom:1px #EFf3ff solid; height:20px; padding:10px 10px 0px 10px;" colspan=2><h3>Eksport og Print</h3></td>
    </tr>
    <tr>
    <td valign=top style="padding:10px 10px 10px 10px;">
	<%
	end function
	
	function filteros09(ptop,pleft,pwdt,txt,visning,tdheight)
	
	select case visning
	case 1
	bgt = "#8CAAE6"
	bgt_border = "#5582d2"
	bgtd = ""
	divbg = "#EFF3FF" 
	case 2 '** print
	bgt = "#cccccc"
	bgt_border = "#999999"
	divbg = "#EFF3FF" 
	case 3
	bgt = "#d6dff5" 'C1D9F0
	bgt_border = "#8CAAE6"
	divbg = "#EFF3FF" 
	case 4
	bgt = "#Cccccc"
	bgt_border = "#999999" 
	divbg = "#FFFFFF"
	
	case 5
	end select
	
	%>
	<div id=Div3 class="jqcorner" style="position:relative; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; border:1px #cccccc solid; overflow:auto; height:<%=tdheight%>px; background-color:<%=divbg%>">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr>
    <td bgcolor="<%=bgt %>" style="border-bottom:1px <%=bgt_border %> solid; padding:0px 10px 0px 10px;" colspan=2><h4><%=txt %></h4></td>
    </tr>
    <tr>
    <td valign=top style="padding:10px 10px 10px 10px;">
	<%
	end function
	
	
	function filterheader(ptop,pleft,pwdt,pTxt)
	
	pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="filter" class="jqcorner" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=pwdt %>px; border:1px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#EFF3FF" style="border-bottom:1px #D6DfF5 solid;">
        <img src="../ill/find.png" /></td>
        <td bgcolor="#EFF3FF" align=left style="border-bottom:1px #D6DfF5 solid;"><b><%=pTxt%>:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function


    function filterheader_2013(ptop,pleft,pwdt,pTxt)
	
	
	
	%>
	<div id="filter" style="position:relative; background-color:#ffffff; padding:10px 10px 10px 10px; width:<%=pwdt %>px; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible; z-index:100;">
    
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF"><h4><%=pTxt%></h4>
	
    
	
	<%
	end function
   

    function filterheader2013(ptop,pleft,pwdt,pTxt)
	
    if pwdt <> 0 then
    pwdt = pwdt&"px"
    else
	pwdt = "100%"
    end if
        'pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="Div1" class="jqcorner" style="position:relative; background-color:#ffffff; padding:10px 10px 10px 10px; width:<%=pwdt%>; border:0px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr>
        <td><h4><%=pTxt%>:</h4></td></tr>
	<tr>
	</table>
        <br /><br />
	<table cellpadding=0 cellspacing=0 border=0 width=60%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function

   
    function filterheaderid(ptop,pleft,pwdt,pTxt,fiVzb,fiDsp,fid,abrel)
	
    '**** Søgefilter expand, Jobbanken

  
    if fid = "filterTreg" then
    pTxt = "<img src=""ill/minus2.gif"" border=0> "& tsa_txt_384 &""
    else
	pTxt = replace(global_txt_119, "|", "&")
	end if
	
	%>
     
	<div id="d_<%=fid %>" style="position:<%=abrel%>; background-color:#ffffff; padding:5px 5px 5px 5px; width:<%=pwdt-10 %>px; border:0px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:<%=fiVzb%>; display:<%=vzDsp%>;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <% if fid = "filterTreg" then %>
    <tr>
       <td><a href="#" id="a_<%=fid %>"><%=pTxt%></a>
         <br /><br />&nbsp;</td></tr>
	<tr>
    <% else %>
     <tr>
       
        &nbsp;</td>
        <td align=left><b><%=pTxt%>:</b></td></tr>
	<tr>
    <%end if%>
	
    </table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	<tr><td bgcolor="#FFFFFF" style="padding:5px; visibility:visible; display:;" id="t_<%=fid %>">
	
    
	
	<%
	end function   
     
     
    function tableDiv(tTop,tLeft,tWdth)
	
	if print = "j" OR media = "print" Or request("print") = "j" then
	bd = 0
	else
	bd = 0
	end if

    if instr(tWdth, "%") > 0 then
        tWdth = tWdth
    else
        tWdth = tWdth&"px"
    end if
	
	%>
	<div id="tableD" style="position:relative; background-color:#ffffff; padding:15px; width:<%=tWdth%>; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:visible;">
    <%
	end function


    function tableDiv_j(tTop,tLeft,tWdth)
	
	if print = "j" OR media = "print" Or request("print") = "j" then
	bd = 0
	else
	bd = 0
	end if

    if instr(tWdth, "%") > 0 then
        tWdth = tWdth
    else
        tWdth = tWdth&"px"
    end if
	
	%>
	<div id="tableD_j" style="position:relative; background-color:#ffffff; padding:3px; width:<%=tWdth%>; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:visible;">
    <%
	end function

  
	
	 function tableDivAbs(tTop,tLeft,tWdth,tHgt,tId, tVzb, tDsp, tZindex)
	
		if print = "j" OR media = "print" Or request("print") = "j" then
	bd = 0
	else
	bd = 0
	end if
	
	%>
	<div id="<%=tId%>" style="position:absolute; background-color:#ffffff; height:<%=tHgt%>px; padding:3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:<%=tVzb%>; display:<%=tDsp%>; z-index:<%=tZindex%>; overflow:auto;">
    <%
	end function
	
	
	
	 function tableDivWid(tTop,tLeft,tWdth,tId, tVzb, tDsp)
	 if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	 
	%>
	<div id="<%=tId%>" style="position:relative; background-color:#ffffff; padding:15px; width:<%=tWdth%>px; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:<%=tVzb%>; display:<%=tDsp%>; z-index:100; border:0px #999999 solid;">
    <%
	end function
    
     
     
    function sideinfo(itop,ileft,iwdt)
	
        iwdt = "100%"
        %>
	<div id="sideinfo" class="jqcorner" style="position:relative; padding:0px 10px 10px 10px; width:<%=iWdt %>; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:visible; z-index:1;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr>
        <td colspan="2" align=left style="border-bottom:0px #C4c4c4 solid;"><h4>Hjælp & Sideinfo:</h4></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td style="padding:5px;">
	
    
	
	<%
	end function
	
	
	 function sideinfoId(itop,ileft,iwdt,ihgt,iId,idsp,ivzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
	%>
	
	<script>
	
	$(document).ready(function() {
    
    
    $("#showpagehelp").click(function() {
    
    //alert("her")
    
    $("#pagehelp_bread").show("fast", function() {
        // use callee so don't have to name the function
    //$(this).next().show("fast", arguments.callee);

    $("#pagehelp_bread").css("display", "");
    $("#pagehelp_bread").css("visibility", "visible");
        
    });
	
	

	$("#pagehelp").hide("fast", function() {
	    // use callee so don't have to name the function
	    //$(this).next().show("fast", arguments.callee);
	});


});
    
    

    $("#hidepagehelp").click(function() {


    $("#pagehelp").show("slow", function() {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });


        $("#pagehelp_bread").hide("slow", function() {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });

     });   
	    
	    
        
    });



    
     

</script>
	
	<div id="<%=iId %>" style="position:absolute; background-color:#ffffff; padding:1px 1px 0px 1px; width:<%=iWdt%>px; border:1px silver solid; border-bottom:0px; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:9000000; overflow:hidden;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center>
        <a href="#" id="showpagehelp" class=alt><%=replace(tsa_txt_390, "|", "&") %> +</a></td>
     </tr>
    </table>
	</div>
	
	
	<div id="<%=ibId %>" style="position:absolute; background-color:#ffffff; padding:5px 5px 5px 5px; width:<%=ibWdt %>px; height:<%=ibhgt %>px; border:1px silver solid; left:<%=ibleft %>px; top:<%=ibtop %>px; visibility:hidden; display:none; z-index:9000000; overflow:auto;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center width=40 style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/lifebelt.png" /></td>
        <td align=left style="border-bottom:1px #C4c4c4 solid;" class=alt><b><%=replace(tsa_txt_390, "|", "&") %></b></td>
         <td style="padding-right:20px;" align="right"><a href="#" id="hidepagehelp" class=alt>[x]</a></td>
       
        </tr>
        <!-- onclick="dsppagehelp('<%=iId%>')" -->
	<tr>
	</table>
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFff" style="padding:5px;">
	
    
	
	<%
	end function
	
    function sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
	%>
	
	<script>
	    function sidemsgclose(idthis) {
	        //alert(idthis)
	        document.getElementById(idthis).style.visibility = "hidden"
	        document.getElementById(idthis).style.display = "none"
	    }
	
	
	</script>
	
	 <div id="<%=iId %>" style="position:absolute; background-color:#FFFFFF; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:10px #CCCCCC solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FF6666" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/ikon_message_24.png" /></td>
        <td bgcolor="#FF6666" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Meddelelse</b></td>
	    <td bgcolor="#FF6666" align=right style="border-bottom:1px #C4c4c4 solid; padding-right:5px;"><a href="#" onclick="sidemsgclose('<%=iId%>')" class="vmenu">[x]</a></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function
	
	  function sidemsgId2(itop,ileft,iwdt,iId,idsp,ivzb)
	%>
	
	<script>
	    function sidemsgclose2(idthis) {
	        //alert(idthis)
	        document.getElementById(idthis).style.visibility = "hidden"
	        document.getElementById(idthis).style.display = "none"
	    }


	
	
	</script>
	
	<div id="<%=iId %>" style="position:absolute; background-color:#FFFFFF; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:2px #c4c4c4 solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FFFF99" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/ikon_message_24.png" /></td>
        <td bgcolor="#FFFF99" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Meddelelse</b></td>
	    <td bgcolor="#FFFF99" align=right style="border-bottom:1px #C4c4c4 solid; padding-right:5px;"><a href="#" onclick="sidemsgclose2('<%=iId%>')" class="vmenu">[x]</a></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" valing="top" style="padding:5px;">
	
    
	
	<%
	end function
	
	function opretNy(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="padding:1px 0px 0px 10px;">
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	function opretNyAB(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:absolute; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="padding:1px 0px 0px 10px;">
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	
	
	function opretNy_blank(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="padding:1px 0px 0px 10px;">
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_blank"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_blank"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	function opretNyJava(url, text, otoppx, oleftpx, owdtpx, java)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="border-bottom:0px; padding:1px 0px 0px 10px;">
	<a href='<%=url %>' onclick="<%=java%>" class='vmenu' alt="<%=text %>"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px; border-bottom:0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" onclick="<%=java %>"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	

     function opretNy_2013(nWdt, nTxt, nLnk, nTgt)%>
	 <div style="background-color:forestgreen; padding:5px 5px 5px 5px; width:<%=nWdt%>px;"><a href="<%=nLnk %>" class="alt" target="<%=nTgt%>"><%=nTxt %> +</a></div>
    <%end function

         function opretNy_2013_java(nWdt, nTxt, nLnk, nTgt, nJav)%>
	 <div style="background-color:forestgreen; padding:5px 5px 5px 5px; width:<%=nWdt%>px;"><a href='<%=nLnk%>' onclick="<%=nJav%>" class='alt'><%=nTxt %></a></div>
    <%end function
	
	
	
	
	
	function insertDelhist(deltype, delid, delnr, delnavn, mid, mnavn)
	
	
	strSQLdelhist = "INSERT INTO delete_hist (deltype, delid, delnr, delnavn, mid, mnavn) VALUES "_
	&" ('"& deltype &"', "& delid &", '"& delnr &"', '"& delnavn &"', "& mid &", '"& mnavn &"')"
	
    'Response.Write strSQLdelhist

	oConn.execute(strSQLdelhist)
	
	end function
	
	
	
	
	
	
	
	
	
	public htmlparseTxt
	function htmlreplace(HTMLstring)
	
	
	    txtBRok = replace(HTMLstring, "<br>", "[#br#]") 
        txtBRok = replace(txtBRok, "<br />", "[#br#]") 
        txtBRok = replace(txtBRok, "<br/>", "[#br#]") 
        txtBRok = replace(txtBRok, "<b>", "[#b#]")
        txtBRok = replace(txtBRok, "</b>", "[#/b#]")
        
        txtBRok = replace(txtBRok, "<p>", "")
        txtBRok = replace(txtBRok, "</p>", "[#br#]")
        
        
        '** Skal lave linjeskift pga. job - print - tilbud
        txtBRok = replace(txtBRok, "</tr>", "[#br#]")

        '** Skal lave linjeskift pga. job - print - tilbud
        txtBRok = replace(txtBRok, "<div>", "[#br#]")
        txtBRok = replace(txtBRok, "</div>", "[#br#]")
       
        txtBRok = replace(txtBRok, "<strong>", "[#strong#]")
        txtBRok = replace(txtBRok, "</strong>", "[#/strong#]")
       
        txtBRok = replace(txtBRok, "<i>", "[#i#]")
        txtBRok = replace(txtBRok, "</i>", "[#/i#]")
        
        txtBRok = replace(txtBRok, "<em>", "[#em#]")
        txtBRok = replace(txtBRok, "</em>", "[#/em#]")
       
        txtBRok = replace(txtBRok, "<u>", "[#u#]")
        txtBRok = replace(txtBRok, "</u>", "[#/u#]")

        txtBRok = replace(txtBRok, "<img", "[#img")
        txtBRok = replace(txtBRok, "</img>", "[#/img#]")
       
        
        HTMLstring = txtBRok
        
        
                Set RegularExpressionObject = New RegExp

                With RegularExpressionObject
                .Pattern = "<[^>]+>"
                .IgnoreCase = True
                .Global = True
                End With

                stripHTMLtags = RegularExpressionObject.Replace(HTMLstring, "")
                htmlparseTxt = replace(stripHTMLtags, "[#", "<")
                htmlparseTxt = replace(htmlparseTxt, "#]", ">")
                Set RegularExpressionObject = nothing
    	
    end function
    
 
 
  
 
 
  public assci_formatTxt
 function assci_format(assci_str)
            
            assci_str = replace(assci_str, "&#248;", "ø")
            assci_str = replace(assci_str, "&#230;", "æ")
            assci_str = replace(assci_str, "&#229;", "å")
            assci_str = replace(assci_str, "&#216;", "Ø")
            assci_str = replace(assci_str, "&#198;", "Æ")
            assci_str = replace(assci_str, "&#197;", "Å")
            assci_str = replace(assci_str, "&#214;", "Ö")
            assci_str = replace(assci_str, "&#246;", "ö")
            assci_str = replace(assci_str, "&#220;", "Ü")
            assci_str = replace(assci_str, "&#252;", "ü")
            assci_str = replace(assci_str, "&#196;", "Ä")
            assci_str = replace(assci_str, "&#228;", "ä")
            
            
            assci_formatTxt = assci_str
 
 end function

    
 public utf_formatTxt
 function utf_format(utf_str)
            
            utf_str = replace(utf_str, "ø", "&#248;")
            utf_str = replace(utf_str, "æ", "&#230;")
            utf_str = replace(utf_str, "å", "&#229;")
            utf_str = replace(utf_str, "Ø", "&#216;")
            utf_str = replace(utf_str, "Æ", "&#198;")
            utf_str = replace(utf_str, "Å", "&#197;")
            utf_str = replace(utf_str, "Ö", "&#214;")
            utf_str = replace(utf_str, "ö", "&#246;")
            utf_str = replace(utf_str, "Ü", "&#220;")
            utf_str = replace(utf_str, "ü", "&#252;")
            utf_str = replace(utf_str, "Ä", "&#196;")
            utf_str = replace(utf_str, "ä", "&#228;")
            
            
            utf_formatTxt = utf_str
 
 end function
 
 public jq_formatTxt
 function jq_format(jq_str)
            

            if len(trim(jq_str)) <> 0 then

            jq_str = replace(jq_str, "ø", "&oslash;")
            jq_str = replace(jq_str, "æ", "&aelig;")
            jq_str = replace(jq_str, "å", "&aring;")
            jq_str = replace(jq_str, "Ø", "&Oslash;")
            jq_str = replace(jq_str, "Æ", "&AElig;")
            jq_str = replace(jq_str, "Å", "&Aring;")
            jq_str = replace(jq_str, "Ö", "&Ouml;")
            jq_str = replace(jq_str, "ö", "&ouml;")
            jq_str = replace(jq_str, "Ü", "&Uuml;")
            jq_str = replace(jq_str, "ü", "&uuml;")
            jq_str = replace(jq_str, "Ä", "&Auml;")
            jq_str = replace(jq_str, "ä", "&auml;")
            jq_str = replace(jq_str, "é", "&eacute;")
            jq_str = replace(jq_str, "É", "&Eacute;")
            jq_str = replace(jq_str, "á", "&aacute;")
            jq_str = replace(jq_str, "Á", "&Aacute;")
            
            jq_formatTxt = jq_str

            else

            jq_formatTxt = ""

            end if
 
 end function    
    
 public htmlparseCSVtxt
 Function htmlparseCSV(HTMLstring)
    
    
        Set RegularExpressionObject = New RegExp

        With RegularExpressionObject
        .Pattern = "<[^>]+>"
        .IgnoreCase = True
        .Global = True
        End With

        stripHTMLtags = RegularExpressionObject.Replace(HTMLstring, "")
        htmlparseCSVtxt = stripHTMLtags

       
       

        Set RegularExpressionObject = nothing

End Function






function opdaterFeriePl(level, highV)

                '**** Ændrer status på planlagt ferie til afholdt ferie ***'
		        '*** (indlæser ny registrering så historik beholdes) ***'
		        '**** Tjekker om der findes reg. i forvejen *****'
		        '11 Ferie Planlagt
		        '14 Ferie Afholdt
                '19 Ferie afholdt uden løn
		        
		        '18 Ferie Fridage Planlagt
		        '13 Ferie fridage brugt


		        
                '26 Aldersreduktion PL
                
                '120 Omsorg 2 PL
                '121 Omsorg 10 PL
                '122 Omsorg K PL



		        '*** Opdaterer for alle hvis det er en admin bruger der logger på ***'
		        
		        if level = 1 then
		        
		        for f = 1 to highV
		        
		        select case f 
                case 1
		        planlagtVal = 11
		        afholdtVal = 14
		        case 2
		        planlagtVal = 18
		        afholdtVal = 13
                case 3
		        planlagtVal = 26
		        afholdtVal = 28
                case 4
		        planlagtVal = 120
		        afholdtVal = 23
                case 5
		        planlagtVal = 121
		        afholdtVal = 10
                case 6
		        planlagtVal = 122
		        afholdtVal = 115
		        end select
		        
		        aktid = 0
		        LoginDato = year(now)&"/"& month(now)&"/"&day(now)
		        
		        '** Finder navn og id på afholdt akt. DVS afholdt ferie / feriefri / Alders reduktion mv ***'
		        strSQLfeafn = "SELECT a.id, a.navn FROM job j"_
		        &" LEFT JOIN aktiviteter a ON (a.fakturerbar = "& afholdtVal &" AND a.aktstatus = 1 AND a.job = j.id) "_
		        &" WHERE j.jobstatus = 1 AND a.id <> 'NULL' GROUP BY a.id ORDER BY a.id DESC"
		        
		    
		        oRec3.open strSQLfeafn, oCOnn, 3
		        if not oRec3.EOF then
		        
		        if oRec3("id") <> "" then
		        aktid = oRec3("id")
		        aktnavn = oRec3("navn")
		        end if
		        
		        end if
		        oRec3.close
		        
		        'Response.Write "<br>aktid" & aktid & "<br>"
		        
		        if cdbl(aktid) <> 0 then
		        
		        
		        '** Opdater ferie/feriefri for den bruger der logger på - kun aktive aktvitieter ***'
		        strSQLfepl = "SELECT * FROM timer "_
                &" LEFT JOIN aktiviteter AS a ON (a.id = taktivitetid AND a.aktstatus = 1) "_
                &" WHERE tfaktim = "& planlagtVal &" AND tdato BETWEEN '2015-12-01' AND '"& LoginDato &"' AND a.aktstatus = 1 ORDER BY tdato"' AND tmnr = "& session("mid")
		        

                'if lto = "glad" AND session("mid") = 1 then
		        'Response.Write strSQLfepl & "<br>"
		        'Response.flush        
                'end if

		        oRec4.open strSQLfepl, oConn, 3
		        while not oRec4.EOF 
		                
		                indtastningfindes = 0
		                strSQLfeaf = "SELECT timer, tdato FROM timer WHERE tfaktim = "& afholdtVal &" AND tdato = '"& year(oRec4("tdato")) &"/"& month(oRec4("tdato")) & "/"& day(oRec4("tdato")) &"' AND tmnr = "& oRec4("tmnr") 'session("mid")
		                
		                'Response.Write strSQLfeaf & "<br><br>"
		                
		                oRec3.open strSQLfeaf, oCOnn, 3
		                if not oRec3.EOF then
		                '*** ignorer da der allerede finsdes indtastning **'
		                indtastningfindes = 1
		                end if
		                oRec3.close
		                
		                'Response.Write indtastningfindes 
		                
		                
		                
		                if cint(indtastningfindes) = 0 then


                        if isNull(oRec4("timer")) <> true AND len(trim(oRec4("timer"))) <> 0 then
                        timerThis = oRec4("timer") 'replace(oRec4("timer"), ",", ".")
                        else
                        timerThis = 0
                        end if

                        if isNull(oRec4("TimePris")) <> true AND len(trim(oRec4("timePris"))) <> 0 then
                        tpThis = replace(oRec4("TimePris"), ",", ".")
                        else
                        tpThis = 0
                        end if

                        if isNull(oRec4("kostpris")) <> true AND len(trim(oRec4("kostpris"))) <> 0 then
                        kpThis = replace(oRec4("kostpris"), ",", ".")
                        else
                        kpThis = 0
                        end if

                        if isNull(oRec4("kurs")) <> true AND len(trim(oRec4("kurs"))) <> 0 then
                        kursThis = replace(oRec4("kurs"), ",", ".")
                        else
                        kursThis = 0
                        end if

                        if isNull(oRec4("timerkom")) <> true AND len(trim(oRec4("timerkom"))) <> 0 then
                        kommThis = replace(oRec4("timerkom"), "'", "''")
                        else
                        kommThis = ""
                        end if






                        if f = 1 then 'ferie afholdt / ferie afholdt uden løn

                          

                            if datePart("m", LoginDato, 2,2) <= 4 then

                                 ferieaarST = year(dateAdd("yyyy", -1, LoginDato)) &"-5-1" 
                                 ferieaarSL = year(LoginDato) &"-4-30"

                            else

                                ferieaarST = year(LoginDato) &"-5-1" 
                                ferieaarSL = year(dateAdd("yyyy", 1, LoginDato)) &"-4-30"

                            end if



                            '** Ferie optjent
                            fe_optjent = 0  
                            strSQLfo = "SELECT COALESCE(SUM(timer), 0) AS fe_optjent FROM timer WHERE tdato between '"& ferieaarST &"' AND '"& ferieaarSL &"' AND tmnr = "& oRec4("tmnr") &" AND (tfaktim = 15 OR tfaktim = 111)" 
                            
                            'response.write "strSQLfo: " & strSQLfo & "<br>"

          
		                    oRec3.open strSQLfo, oCOnn, 3
		                    if not oRec3.EOF then

                            fe_optjent = oRec3("fe_optjent")

                            end if
                            oRec3.close
                
                              


                            'response.Write "<br>fe_optjent: " & fe_optjent & "<br>"

                            '** Ferie afholdt i ferieåret 
                            fe_afholdtmlon = 0
                            strSQLfa = "SELECT COALESCE(SUM(timer), 0) AS fe_afholdtmlon FROM timer WHERE tdato between '"& ferieaarST &"' AND '"& ferieaarSL &"' AND tmnr = "& oRec4("tmnr") &" AND (tfaktim = 14)" 

                            'response.write "<br><br>strSQLfa: " & strSQLfa & "<br>"

                            oRec3.open strSQLfa, oCOnn, 3
		                    if not oRec3.EOF then

                            fe_afholdtmlon = oRec3("fe_afholdtmlon")

                            end if
                            oRec3.close
                        


                            

                            ''if lto = "wwf" AND session("mid") = 1 then
                            'response.write "fe_optjent: "& fe_optjent
                            'response.Write "<br>fe_afholdtmlon: " & fe_afholdtmlon & "<br>"
                            'response.write "timerthis" & timerthis
                            'response.write "<br>afholdetUlon: " & (fe_optjent - (fe_afholdtmlon + timerThis)) * -1 & "<br><br>"
                            
                            'end if


                            '*** Overfører til afholdt uden løn
                            if (fe_afholdtmlon + timerThis) >= fe_optjent AND fe_optjent > 0 then


                          

                            afholdetUlon = (fe_optjent - (fe_afholdtmlon + timerThis)) * -1  

                            timerThisOpr = timerThis
                            timerThis = afholdetUlon
                            afholdtValOpr = afholdtVal
                            afholdtVal = 19
                            aktidOpr = aktid
		                    aktnavnOpr = aktnavn 


                                    '** Finder navn og id på afholdt ferie akt. U LØN  ***'
		                            strSQLfeafn = "SELECT a.id, a.navn FROM job j"_
		                            &" LEFT JOIN aktiviteter a ON (a.fakturerbar = 19 AND a.aktstatus = 1 AND a.job = j.id) "_
		                            &" WHERE j.jobstatus = 1 AND a.id <> 'NULL' GROUP BY a.id ORDER BY a.id DESC"
		        
		    
		                            oRec3.open strSQLfeafn, oCOnn, 3
		                            if not oRec3.EOF then
		        
		                            if oRec3("id") <> "" then
		                            aktid = oRec3("id")
		                            aktnavn = oRec3("navn")
		                            end if
		        
		                            end if
		                            oRec3.close


                                call insertAfholdtTimer
                    
                                '*** Sletter den planlagte **'
		                        strSQLdel = "DELETE FROM timer WHERE tid = "& oRec4("tid")
		                        oConn.execute(strSQLdel)


                            '*** Nedskriver timer
                            timerThis = (TimerThisOpr - afholdetUlon)
                            afholdtVal = afholdtValOpr

                            aktid = aktidOpr
		                    aktnavn = aktnavnOpr
                           


                            end if

                        
                            '** 


                        end if 'f


                        

                                '*** Indlæser / overfører til AFHOLDT TYPE ***'
                                if (f > 0 AND timerThis > 0) OR (f = 1 AND timerThis > 0 AND (timerThis > afholdetUlon)) then


		                                call insertAfholdtTimer
		                
                       
		                        timerThis = 0
		                
		                        '*** Sletter den planlagte **'
		                        strSQLdel = "DELETE FROM timer WHERE tid = "& oRec4("tid")
		                        oConn.execute(strSQLdel)

                        
		                        end if' timerThis > 0                 


		                end if
		        
		                
		        
		        
		        oRec4.movenext
		        wend
		        oRec4.close
                
                end if '*** aktid <> 0
                
                next
                
                end if 'level
                
                'if lto = "wwf" AND session("mid") = 1 then
                'Response.end
                'end if

end function



public afholdtVal, kommThis, aktid, aktnavn, kursThis , tpThis, kpthis
sub insertAfholdtTimer
        


                timerThis = replace(timerThis, ",", ".")

                stTidThis = formatdatetime(oRec4("sttid"), 3)
                slTidThis = formatdatetime(oRec4("sltid"), 3)
        
                strSQLfeins = "INSERT INTO timer "_
		        &"("_
		        &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		        &" timerkom, TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		        &" editor, kostpris, offentlig, seraft, godkendtstatus, "_
		        &" godkendtstatusaf, sttid, sltid, valuta, kurs "_
		        &") "_
		        &" VALUES "_
		        &" (" _
		        & timerThis &", "& afholdtVal &", "_
		        & "'"& year(oRec4("tdato")) &"/"& month(oRec4("tdato")) & "/"& day(oRec4("tdato")) &"', "_
		        & "'"& oRec4("tmnavn") &"', "_
		        & oRec4("tmnr") &", "_
		        & "'"& oRec4("tjobnavn") &"', "_
		        & "'"& oRec4("tjobnr") &"', "_
		        & "'"& oRec4("tknavn") &"', "_
		        & oRec4("tknr") &", "_
		        & "'"& kommThis &"', "_
		        & aktid &", "_
		        & "'"& aktnavn &"', "_
		        & oRec4("Taar") &", "_
		        & tpThis &", "_
		        & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', "_
		        & oRec4("fastpris") &", "_
		        & "'"& time &"', "_
		        & "'"& oRec4("editor") &"', "_
		        & kpThis &", "_
		        & oRec4("offentlig") &", "_
		        & oRec4("seraft") &", "_
		        & oRec4("godkendtstatus") &", "_
		        & "'"& oRec4("godkendtstatusaf") &"', "_
		        & "'"& stTidThis &"', "_
                & "'"& slTidThis &"', "_
		        & oRec4("valuta") &", "_
		        & kursThis &")"


            
                'if lto = "glad" AND session("mid") = 1 then
                'response.write strSQLfeins &"<br><br>"
                'response.flush
                'end if

                oConn.execute(strSQLfeins)



end sub


public SY_usejoborakt_tp, SY_fastpris, SY_laasmedtpbudget
function skaljobSync(jobid)

            '** Skal job sync ****'
			SY_usejoborakt_tp = 0
		    SY_laasmedtpbudget = 0
        	SY_fastpris = 0

			strSQLj = "SELECT jobnavn, usejoborakt_tp, fastpris, laasmedtpbudget FROM job WHERE id = " & jobid
			
			'Response.Write strSQLj
			'Response.end
			oRec4.open strSQLj, oConn, 3
			if not oRec4.EOF then
			
			SY_usejoborakt_tp = oRec4("usejoborakt_tp")
			SY_fastpris = oRec4("fastpris")
            SY_laasmedtpbudget = oRec4("laasmedtpbudget")
			
			end if
			oRec4.close
			
			

end function

function syncJob(jobid)

                call akttyper2009(2)
                 			
			                strSQLaktSum = "SELECT SUM(budgettimer) sumakttimer, fakturerbar, SUM(aktbudgetsum) AS sumaktbudget FROM aktiviteter "_
			                &" WHERE job =  "& jobid & " AND("& aty_sql_fakbar &") AND aktfavorit = 0 GROUP BY job"
			                oRec2.open strSQLaktSum, oConn, 3
			                if not oRec2.EOF then
                			
			                sumakttimer = replace(oRec2("sumakttimer"), ",", ".")
			                sumaktbudget = replace(oRec2("sumaktbudget"), ",", ".")
                			
			                end if
			                oRec2.close
				
				strSQLsync = "UPDATE Job SET budgettimer = "& sumakttimer &", "_
				&" ikkebudgettimer = 0, jobtpris = "& sumaktbudget &" WHERE id = "& jobid
				
				oConn.execute(strSQLsync)
				
				'Response.Write strSQLsync
				'Response.Write " -- her"
				'Response.end

end function


'option explicit 

' Simple functions to convert the first 256 characters 
' of the Windows character set from and to UTF-8.

' Written by Hans Kalle for Fisz
' http://www.fisz.nl

'IsValidUTF8
'  Tells if the string is valid UTF-8 encoded
'Returns:
'  true (valid UTF-8)
'  false (invalid UTF-8 or not UTF-8 encoded string)
function IsValidUTF8(s)
  dim i
  dim c
  dim n

  IsValidUTF8 = false
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      select case n
      case 1
        exit function
      case 2
        if (c and &HE0) <> &HC0 then
          exit function
        end if
      case 3
        if (c and &HF0) <> &HE0 then
          exit function
        end if
      case 4
        if (c and &HF8) <> &HF0 then
          exit function
        end if
      case else
        exit function
      end select
      i = i + n
    else
      i = i + 1
    end if
  loop
  IsValidUTF8 = true 
end function

'DecodeUTF8
'  Decodes a UTF-8 string to the Windows character set
'  Non-convertable characters are replace by an upside
'  down question mark.
'Returns:
'  A Windows string
function DecodeUTF8(s)
  dim i
  dim c
  dim n

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191 
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s 
end function

'EncodeUTF8
'  Encodes a Windows string in UTF-8
'Returns:
'  A UTF-8 encoded string
'function EncodeUTF8(s)
'  dim i
'  dim c

'  i = 1
'  do while i <= len(s)
'    c = asc(mid(s,i,1))
'    if c >= &H80 then
'      s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
'      i = i + 1
'    end if
'    i = i + 1
'  loop
'  EncodeUTF8 = s 
'end function


Function EncodeUTF8(s)
        Dim i
        Dim c
       
        i = 1
        Do While i <= len(s)
            c = asc(mid(s, i, 1))
            If c >= &H80 Then
                s = left(s, i - 1) + chr(&HC2 + ((c And &H40) / &H40)) + chr(c And &HBF) + mid(s, i + 1)
                i = i + 1
            End If
            i = i + 1
        Loop
        
        's = Replace(s, "Ã,", "&oslash;")
        EncodeUTF8 = s
End Function


public strDay_30
function dato_30(dagDato, mdDato, aarDato)

                if dagDato > 28 then 
				select case mdDato
				case "2"
				    
				    if len(trim(aarDato)) = 2 then
				    aarDato = "20" & aarDato
				    else
				    aarDato = aarDato
				    end if
				    
				    select case aarDato
				    case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
				    strDay_30 = 29
				    case else
				    strDay_30 = 28
				    end select
				    
				case "4", "6", "9", "11"
				    if dagDato > 30 then
				    strDay_30 = 30
				    else
				    strDay_30 = dagDato
				    end if
				case else
				strDay_30 = dagDato
				end select
				else
				strDay_30 = dagDato
				end if

end function
%>





  
 
 