

<%
if len(trim(request("menu"))) <> 0 then
menu = request("menu")
else
menu = menu 'sat i fil
end if

topmenuSel = menu


'****** variable til emner. status og medarb sættes her *****
if menu = "crm" then
	if len(request("medarb")) <> 0 then 
	medarb = request("medarb")
	else
	medarb = session("mid")
	end if
	if len(request("emner")) <> 0 then 
	emner = request("emner")
	else
	emner = 0
	end if
	
	if len(request("status")) <> 0 then 
	status = request("status")
	else
	status = 0
	end if
	
	if len(request("id")) = 0 then
	id = 0
	else
	id = request("id")
	end if
end if
'**************************************************************'





level = session("rettigheder")
%>

<div name="about" id="about" style="position:absolute; top:52px; left:400px; padding:20px; visibility:hidden; background-color:#FFFFE1; border:2 orange dashed; padding:10px; z-index:2000;">
<table height="115" width="275" cellpadding=0 cellspacing=0 border=0>
<%
strSQL = "SELECT * FROM licens WHERE id = 1"
oRec.open strSQL, oConn, 3
if not oRec.EOF then
licenshaver = oRec("licens")
%>
<tr><td></td><td align="right"><a href="#" onclick="hideabout()" class="red">[x]</a></td></tr>
<tr><td width="75">Firma:</td><td><b><%=oRec("licens")%></b></td></tr>
<tr><td>Licensnr:</td><td><b><%=oRec("key")%></b></td></tr>

<tr><td>DB version:</td><td> 
<%
strSQL2 = "SELECT dbversion FROM dbversion ORDER BY id DESC"
oRec2.open strSQL2, oConn, 3
if not oRec2.EOF then
Response.write oRec2("dbversion")
end if
oRec2.close
%>
</td></tr>

<tr><td>Licenstype: </td><td><%=oRec("licenstype")%></td></tr>
<tr><td>Brugere:</td><td> <%=oRec("klienter")%></td></tr>
<tr><td>Support: </td><td><a href="mailto:support@outzource.dk" class=vmenu>support@outzource.dk</a></td></tr>
<%
'** tjek om der er tilvalgt CRM ***
'len_licensType = instr(oRec("key"), "B")
'	if len_licensType > 0 then
	licensType = "CRM"
'	end if
end if
oRec.close
%>
</table>
</div>

<%


'*** Speciel tilrettet kundedesign ***
'Response.write thisfile
if thisfile = "joblog_k" OR (thisfile = "infobase" AND kview = "j") OR (thisfile = "sdsk" AND kview = "j") then
	
	select case lto
	case "kringit"
	tablebg = "#C5161C" 
	firmalogo = "<img src='../ill/kringit/kring_value.gif' alt='' border='0'>"
	keys = "<img src='../ill/blank.gif' alt='' border='0'>"
	loginfile = "login_kunder.asp"
	%>
	<div name="waterm" id="waterm" style="position:absolute; top:142; left:0; z-index:-1000;">
	<img src='../ill/kringit/kring_waterm_st.gif' alt='' border='0'>
	</div>
	<% 
	case else
	tablebg = "#003399"
	firmalogo = "<img src='../ill/logo_topbar.gif' alt='' border='0'>"
	keys = "<img src='../ill/keys.gif' alt='' border='0'>"
	loginfile = "login_kunder.asp"
	
	end select

else
tablebg = "#3B5998" '"#003399"
firmalogo = "<img src='../ill/logo_topbar_2008_rus.gif' alt='' border='0'>"
keys = "<img src='../ill/keys.gif' alt='' border='0'>"
loginfile = "login.asp"
end if
%>


<table cellspacing="0" cellpadding="0" border="0" bgcolor="<%=tablebg%>" width=100%>
<tr>
<td valign=top><a href="../<%=loginfile%>" target="_top"><%=firmalogo%></a></td>
<td align="right" valign=bottom>
	
	
	<table cellpadding="2" cellspacing="0" border=0>
	<tr>
	<td valign="top"><%=keys%></td>
	<td>
	<font style="color:#ffffff;font-family:arial,sans-serif;font-size:11px;">Du er <strong>logget ind</strong> som:
		&nbsp;<%=session("user")%>&nbsp;&nbsp;<font color="#FFCC00">(<%=licenshaver%>)</font> 
		<%
		strSQL = "SELECT lastlogin FROM medarbejdere WHERE Mnavn ='" & session("user") &"'"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			strDato = oRec("lastlogin")
			end if
			oRec.close
		%>
		&nbsp;&nbsp;<br>Dato:&nbsp;<%=strDato%> - Sidste login:&nbsp;<%=session("strLastlogin")%></font>
	</td>
	<td valign=bottom style="padding-bottom:0px; padding-right:5px;">
	<%if session("stempelur") <> 0 then %>
	<a href="stempelur.asp?func=redloginhist&medarbSel=<%=session("mid")%>&showonlyone=1&hidemenu=1&id=0&rdir=sesaba" target="_top"><img src="../ill/logud.gif" width="46" height="14" alt="" border="0"></a>
	<%else %>
	<a href="../sesaba.asp" target="_top"><img src="../ill/logud.gif" width="46" height="14" alt="" border="0"></a>
    <%end if %>
	
    </td></tr>
    </table>
	
</td>
</tr>
</table>



<br>
<%



lefthelpinfo = 181
if thisfile = "joblog_k" OR (thisfile = "infobase" AND kview = "j") OR (thisfile = "filarkiv" AND (kundelogin = 1 OR nomenu = 1)) OR (thisfile = "sdsk" AND kview = "j") OR (thisfile = "sdsk_kontrolpanel" AND kview = "j") then 'kunde login side. / infobase kundeview

else
	if level <= 3 AND licensType = "CRM" OR level = 4 AND licensType = "CRM" then%>
	<div id="mainmenu" style="position:absolute; top:14px; left:210px; z-index:1000;">
	
	<table cellspacing="1" cellpadding="0" border="0">
	<tr>
	
	<%
	'*** finder de knapper (on/off) der skal vises ***
	select case menu
	case "crm"
	bgcrm = "#5C75AA"
	bgsdsk = "#3B5998"
	bgtsa = "#3B5998"
	bgerp = "#3B5998"
	case "tok"
	case "eml"
	case "sdsk"
	bgcrm = "#3B5998"
	bgsdsk = "#5C75AA"
	bgtsa = "#3B5998"
	bgerp = "#3B5998"
	case "erp"
	bgcrm = "#3B5998"
    bgerp = "#5C75AA"
    bgtsa = "#3B5998"
    bgsdsk = "#3B5998"
    case else
    bgcrm = "#3B5998"
	bgtsa = "#5C75AA"
	bgerp = "#3B5998"
	bgsdsk = "#3B5998"
    end select
	
	imgtok = "gear.png"
	
	
	'*** Timereg side ***
	treg0206thisMidTopmenu = 1
	strSQL = "SELECT timereg FROM medarbejdere WHERE mid = "& session("mid")
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	treg0206thisMidTopmenu = oRec("timereg") 
	end if
	oRec.close 
	


		if cint(treg0206thisMidTopmenu) = 0 then
			tregLinkTopmenu = "timereg.asp?menu=timereg"
		else
			tregLinkTopmenu = "timereg_2006_fs.asp"
		end if
	%>
	
	
			<td align=center><div id="tsa" style="position:relative; top:0px; padding:4px; left:0px; background-color:<%=bgtsa%>; width:60px; border:0px #5C75AA solid; border-bottom:0px;" onmouseover="bgcolthisON('tsa');" onmouseout="bgcolthisOFF('tsa','<%=bgtsa%>');">
			<a href="<%=tregLinkTopmenu%>" class='modul' target="_top">TSA</a>
			</div>
			</td>
			
			<%
			call erCRMaktiv()
			if cint(crmOnOff) = 1 AND (level <= 4) then%>
			
			
			<td align=center><div id="crm" style="position:relative; top:0px; padding:4px; left:0px; background-color:<%=bgcrm%>; width:60px; border:0px #5C75AA solid; border-bottom:0px;" onmouseover="bgcolthisON('crm');" onmouseout="bgcolthisOFF('crm','<%=bgcrm%>');">
			<a href="crmkalender2.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0" class='modul' target="_top">CRM</a>
			</div>
			</td>
			
			
			<%end if%>
			
			<%
			call erSDSKaktiv()
			if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then%>
			<td align=center><div id="sdsk" style="position:relative; top:0px; padding:4px; left:0px; background-color:<%=bgsdsk%>; width:60px; border:0px #5C75AA solid; border-bottom:0px;" onmouseover="bgcolthisON('sdsk');" onmouseout="bgcolthisOFF('sdsk','<%=bgsdsk%>');">
			<a href="sdsk.asp?menu=sdsk" class='modul' target="_top">SDSK</a>
			</div>
			</td>
			
			<%end if%>
			<%
			call erERPaktiv()
			if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then%>
			<td align=center><div id="erp" style="position:relative; top:0px; padding:4px; left:0px; background-color:<%=bgerp%>; width:60px; border:0px #5C75AA solid; border-bottom:0px;" onmouseover="bgcolthisON('erp');" onmouseout="bgcolthisOFF('erp','<%=bgerp%>');">
			<a href="erp_tilfakturering.asp?menu=erp" class='modul' target="_top">ERP</a>
			</div>
			</td>
			<%end if%>
			<!--<td>&nbsp;<a href="email.asp?menu=eml" class='alt'><img src="../ill/<=imgemail%>" width="78" height="22" alt="" border="0"></a></td>-->
	        
	        <!--
			call erBGTaktiv()
			if cint(bgtOnOff) = 1 AND (level = 1) AND (session("mid") = 73 OR session("mid") = 1 OR session("mid") = 36) OR (cint(bgtOnOff) = 1 AND session("mid") = 20) then%>
			<td valign=top><a href="budget_bruttonetto.asp?menu=bgt" class='alt' target="_top"><img src="../ill/<%=imgbgt%>" width="78" height="22" alt="Budget modul" border="0"></a></td>
			<end i
			-->
	        
	        
	
	<%
	'*** Kun for administratorer ****
	if level = 1 then 
	%>
	<td style="padding:5px 0px 0px 10px;"><a href="kontrolpanel.asp?menu=tok" target="_top"><img src="../ill/<%=imgtok%>" alt="Kontrolpanel" border="0"></a></td>
		
	<% 
	end if
	   %>
	   
	   
	
	<td style="padding:5px 0px 0px 5px;"><a href="#" onclick="showabout()"; class='vmenuglobal'><img src="../ill/information.png" alt="Licens Info" border="0"></a></td>
	<td>&nbsp;</td>
	<td style="padding:5px 0px 0px 3px;"><a href="help.asp?menu=<%=topmenuSel%>" target="_blank"><img src="../ill/lifebelt.png" alt="Hjælp" border="0"></a></td>

	   
	        
	</tr>
	</table>
	</div>
	<%
	end if
%>




<%
end if%>
