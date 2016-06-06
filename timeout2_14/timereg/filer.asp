<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->



<%

'**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_showjob"


        kundeid = request("kundeid")

        strJobs = "<option value='0'>Nej</option>" 

       strSQL = "SELECT jobnavn, id, jobnr FROM job WHERE ((jobstatus = 1 AND fakturerbart = 1) OR jobstatus = 3) AND jobknr = "& kundeid &" ORDER BY jobnavn"
	
        oRec.open strSQL, oConn, 3
        while not oRec.EOF 

        strJobs = strJobs & "<option value='"& oRec("id") &"'>"& oRec("jobnavn") &" ("& oRec("jobnr") &")</option>" 


        oRec.movenext
        wend 
        oRec.close

       
        '*** ÆØÅ **'
        call jq_format(strJobs)
        strJobs = jq_formatTxt
        
        response.write strJobs

        end select
        response.end
        end if


    	sub sideindholdtop%>
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible; display:; z-index:50;">
	<h4>Filarkiv</h4>
	<table cellspacing=1 bgcolor="#ffffff" cellpadding=2 border=0>
	<%end sub
	
	sub sideindholdbund%> 
	</table>
	<br>
	<br>
	&nbsp;
	</div>
	<br><br>&nbsp;
	<%end sub


 %>
<script>
function NewWin(url)    {
window.open(url, 'Help', 'width=700,height=580,scrollbars=1,toolbar=0,menubar=1');    
}

function popUp(URL,width,height,left,top) {
		window.open(URL, 'navn', 'left='+left+',top='+top+',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
}
</script>

<%
if session("user") = "" then

	errortype = 5
	call showError(errortype)
	else
	%>
	
    <SCRIPT src="inc/filer_jav.js"></script>
	<%
	
	id = request("id")

    if len(trim(request("kundeid"))) <> 0 then
	kundeid = request("kundeid")
    else 
	kundeid = 0
    end if

	thisfile = "filarkiv"
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
    if len(request("fms")) <> 0 then
        if request("vistommefoldere") <> 0 then
        vistommefoldere = 1
        vistommefoldereCHK = "CHECKED"
        else
        vistommefoldere = 0
        vistommefoldereCHK = ""
        end if

        response.cookies("TSA")("tommefoldere") = vistommefoldere

    else
        
        if request.cookies("TSA")("tommefoldere") = "1" then
        vistommefoldere = 1
        vistommefoldereCHK = "CHECKED"
        else
        vistommefoldere = 0
        vistommefoldereCHK = ""
        end if

    end if
	
	if request("nomenu") = "1" then
	nomenu = 1
	else
	nomenu = 0
	end if

    if len(trim(request("sortby"))) <> 0 then
    sortby = request("sortby")
    response.Cookies("tsa")("filarkivsortby") = sortby
    else
        if request.Cookies("tsa")("filarkivsortby") <> "" then
        sortBy = request.Cookies("tsa")("filarkivsortby")
        else
        sortby = "navn"
        end if
    end if
	
    if sortby = "navn" then
    filorderby = "fo.navn, filnavn"
    else
    filorderby = "fi.dato DESC"
    end if

	func = request("func")
	'Response.write kundeid
	level = session("rettigheder")
	




	

	
	
	
	select case func
	case "redfil"
       
        intFolderid = 0 
	
	strSQL = "SELECT filnavn, id, adg_kunde, adg_alle, type, folderid FROM filer WHERE id =" & id
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	
	'filnavn = oRec("filnavn")
	adg_kunde = oRec("adg_kunde")
	adg_alle = oRec("adg_alle")
	intType = oRec("type")
	intFolderid = oRec("folderid")
	
	end if
	oRec.close 
	
	'call sideindholdtop


        %>
	<div id="Div1" style="position:absolute; left:20px; top:20px; padding:10px; background-color:#FFFFFF; visibility:visible; display:; z-index:50;">
	
	<%
    oimg = "icon_cabinet_open.png"
	oleft = 0
	otop = 0
	owdt = 400
    oskrift = "Fil - Indstillinger" '"Filarkiv"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
    
     %>
	
	
	 <%
	 
	 tTop = 0
	tLeft = 0
	tWdth = 400
	
	
	call tableDiv(tTop,tLeft,tWdth)

	%><form action="filer.asp?func=dbredfil&id=<%=id%>&kundeid=<%=kundeid%>&jobid=<%=jobid%>&nomenu=<%=nomenu%>" method="post">
        <table cellpadding="2" cellspacing="1" width="100%">
	<tr>
		<td>
		
		<br><b>Filtype</b>: (Dokument, billede ell. firma logo)<br>
		<select name="FM_filtype">
		<%select case intType
		case 1 
		logoSel = "SELECTED"
		offSel = ""
		bilSel = ""
		case 5
		logoSel = ""
		offSel = "SELECTED"
		bilSel = ""
		case 7
		bilSel = "SELECTED"
		logoSel = ""
		offSel = ""
		end select%>
		<option value="5" <%=offSel%>>Dokument</option>
		<option value="7" <%=bilSel%>>Billede (screenshot)</option>
		<option value="1" <%=logoSel%>>Firma Logo</option>
		</select><br><br>
	<b>Folder:</b><br>
	
	
	<%
	strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
	&" fo.id AS foid, fo.kundese, k.kkundenavn, k.kkundenr "_
	&" FROM foldere fo "_
    &" LEFT JOIN kunder AS k ON (k.kid = fo.kundeid) "_
	&" WHERE fo.kundeid = "& kundeid &" OR fo.kundeid = 0 "& jobidSQL &" OR (fo.id = 500) OR (fo.id = "& intFolderid &")  ORDER BY kundeid, foldernavn"
	
	'Response.write strSQL
	'Response.flush
	%>
	<select name="FM_folderid" id="FM_folderid" style="width:300px;">
	<%
	
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	if cint(oRec("foid")) = cint(intFolderid) then
	fSel = "SELECTED"
	else
	fSel = ""
	end if


    if oRec("kundeid") <> 0 then
    kTxt = " - " & oRec("kkundenavn") & " ("& oRec("kkundenr") &")"
    else
    kTxt = ""
    end if
	%>
	<option value="<%=oRec("foid")%>" <%=fSel%>><%=oRec("foldernavn") & " " & kTxt%> </option>
	<%
	oRec.movenext
	wend
	oRec.close %>
	</select>
	
	<br><br>
	<b>Adgangs rettigheder:</b><br>
	Hvilke brugergrupper skal have adgang til dette dokument:<br>
	<%
	if adg_kunde = 1 then
	chkKunde = "CHECKED"
	else
	chkKunde = ""
	end if
	%>
	<input type="checkbox" name="FM_adg_kunde" id="FM_adg_kunde" value="1" <%=chkKunde%>> Kontakter. <br>
    <input id="Checkbox1" type="checkbox" checked disabled /> TimeOut System Administrator(er).<br>
	
	<%
	if adg_alle = 1 then
	chkAlle = "CHECKED"
	else
	chkAlle = ""
	end if
	%>
	<input type="checkbox" name="FM_adg_alle" id="FM_adg_alle" value="1" <%=chkAlle%>> TimeOut Alle brugere. 
	<br><br>
	<input type="submit" value="Opdater >>" />

    </td></tr>
	</table>
        </form>
    </div>        
    <%
	
	
	case "dbredfil"
	
	intType = request("FM_filtype")
	
	if len(request("FM_adg_kunde")) <> 0 then
	adg_kunde = 1
	else
	adg_kunde = 0
	end if
	
	if len(request("FM_adg_alle")) <> 0 then
	adg_alle = 1
	else
	adg_alle = 0
	end if
	
	folderid = request("FM_folderid")


	strSQLupd = "UPDATE filer SET "_
	&" type = "& intType &", folderid = "& folderid &", "_
	&" adg_kunde = "& adg_kunde &", adg_alle = "& adg_alle &""_
	&" WHERE id = "& id
	
	oConn.execute(strSQLupd)
	
	'if cint(nomenu) <> 1 then
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	'else
	'Response.redirect "filer.asp?kundeid="&kundeid&"&jobid="&jobid&"&nomenu="&nomenu 
	'end if
		
	
	
	case "slet"
	'*** Her spørges om det er ok at der slettes en folder ***
	
	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	<%
	
	slttxt = "<b>Du er ved at <b>slette</b> en folder!</b><br />"_
	&" Når en folder slettes slettes alle de filer der ligge i folderen også.<br><br>"
	slturl = "filer.asp?func=sletok&id="&id&"&kundeid="&kundeid&"&jobid="&jobid&"&nomenu="&nomenu
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,90)
	
	
	
	case "sletok"
	
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		
	
	'*** Her slettes folder og filer ***
	strSQL = "SELECT filnavn, id FROM filer WHERE folderid = "& id 
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		
		'Response.write strPath
		'*** Sletter filen fysisk ***
		on Error resume Next 
		strPath =  "d:\dkdomains\outzource\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & oRec("filnavn")
		Set fsoFile = FSO.GetFile(strPath)
		fsoFile.Delete
		
		oConn.execute("DELETE FROM filer WHERE id = "& oRec("id") &"")
		
	oRec.movenext
	wend
	oRec.close
	
	Set FSO = nothing
	
	oConn.execute("DELETE FROM foldere WHERE id = "& id &"")
	
	'if cint(nomenu) <> 1 then 
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	'else
	'Response.redirect "filer.asp?kundeid="&kundeid&"&jobid="&jobid&"&nomenu="&nomenu 
	'end if
	
	
	case "sletfil"
	'*** Her spørges om det er ok at der slettes en fil ***
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<%
	
	slttxt = "<b>Du er ved at <b>slette</b> en fil!</b><br />"_
	&"Er du sikker på at du vil slette denne fil?<br><br>"
	slturl = "filer.asp?func=sletfilok&id="&id&"&kundeid="&kundeid&"&jobid="&jobid&"&nomenu="&nomenu
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,90)
	

	
	case "sletfilok"
	'*** Her slettes en fil ***
	'ktv
	'strPath =  "E:\www\timeout_xp\wwwroot\ver2_1\upload\"&lto&"\" & Request("filnavn")
	'Abusiness
	strPath =  "d:\dkdomains\outzource\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & Request("filnavn")
	'Response.write strPath
	
	on Error resume Next 

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set fsoFile = FSO.GetFile(strPath)
	fsoFile.Delete
	
	Set FSO = nothing
	
	oConn.execute("DELETE FROM filer WHERE id = "& id &"")
	
	'if cint(nomenu) <> 1 then 
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	'else
	'Response.redirect "filer.asp?kundeid="&kundeid&"&jobid="&jobid&"&nomenu="&nomenu 
	'end if
	
	
	case "dbopretfo", "dbredfo"
	
	strNavn = request("FM_navn")
	if len(request("FM_kundese")) then
	kundese = 1
	else
	kundese = 0
	end if
	

        'response.write "kundeid:" & request("kundeid")
        'response.end 

	
	sqldato = year(now) & "/" & month(now) & "/" & day(now)
	'intJobid = request("FM_jobid")
	
	if len(strNavn) <>  0 then
		
		if func = "dbopretfo" then
		strSQL = "INSERT INTO foldere (navn, kundeid, kundese, jobid, editor, dato) VALUES "_
		&" ('"& strNavn &"', "& kundeid &", "& kundese &", "& jobid &", '"& session("user") &"', '"& sqldato &"')"
		else
		strSQL = "UPDATE foldere SET navn = '"& strNavn &"', "_
		&" kundese = "& kundese &", kundeid = "& kundeid &", jobid = "& jobid &", editor = '"& session("user") &"', dato = '"& sqldato &"' WHERE id = " & id
		end if
		
		
		
		oConn.execute(strSQL)
		
		
		
		
		'if cint(nomenu) <> 1 then
		Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")
		'else
		'Response.redirect "filer.asp?kundeid="& kundeid &"&jobid="&jobid&"&nomenu="&nomenu
		'end if
		
	else
	
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		useleftdiv = "j"
		errortype = 8
		call showError(errortype)
	
	end if
	
	case "oprfo", "fored"
		
           

		if func = "fored" then
		
			strSQL = "SELECT navn, id, kundese, jobid, editor, dato FROM foldere WHERE id = "& id
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			
			strNavn = oRec("navn")
			kundese = oRec("kundese")
			intJobid = oRec("jobid")
			strEditor = oRec("editor")
			dtDato = oRec("dato") 
			
			end if
			oRec.close 
			
			dbfunc = "dbredfo"
            sideOskrift = "Rediger"
		
		else
			
			intJobid = jobid
			strNavn = ""
			kundese = 0
			dbfunc = "dbopretfo"
		    sideOskrift = "Opret"

		end if
		
		
		if kundese = "1" then
		ksCHK = "CHECKED"
		else
		ksCHK = ""
		end if
		
		
		
	%>
	<div id="sindhold" style="position:absolute; left:20px; top:20px; padding:10px; background-color:#FFFFFF; visibility:visible; display:; z-index:50;">
	
	<%
    oimg = "icon_cabinet_open.png"
	oleft = 0
	otop = 0
	owdt = 400
    oskrift = sideOskrift &" folder" '"Filarkiv"
	
	call sideoverskrift_2014(oleft, otop, owdt, oskrift)
    
     %>
	
	
	 <%
	 
	 tTop = 0
	tLeft = 0
	tWdth = 400
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	
	 
                          'tdheight = 300
                          'ptop = 0
                          'pleft = 0
                          'pwdt = 400
            
                         'call filteros09(ptop, pleft, pwdt, "Opret folder", 1, tdheight)
                        
                         %>
                                <form action="filer.asp?func=<%=dbfunc%>&id=<%=id%>&nomenu=<%=nomenu%>" method="post">
                         <table cellspacing=0 cellpadding=0 border=0 width=100%>
                  
	
	
	<%if func = "fored" then
	
	if len(dtDato) <> 0 then
	showDato = formatdatetime(dtDato, 1)
	else
	showDato = ""
	end if%>
	
	
	
	
	<tr><td colspan=2><span style="color:#999999;"><i>Sidst redigeret af: <%=strEditor%> d. <%=showDato%></i></span></td></tr>
	<%end if%>
	<tr><td valign=top colspan=2><br /><br /><b>Foldernavn:</b></td>
	</tr>
	<tr><td valign=top colspan=2><input type="text" name="FM_navn" id="FM_navn" style="width:220px;" value="<%=strNavn%>"><br>
	<input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=ksCHK%>> Denne folder skal være tilgængelig for eksterne kontakter (kunder). <br><br>
	</td></tr>
	<tr><td valign=top colspan=2>

        <br /><b>Kunde:</b><br />
        <select name="kundeid" id="kundeid" size="1" style="width:325px;">
		<option value="0">Alle</option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE Kid <> 0 ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>    <br />
<br />


	<b>Skal denne folder være knyttet til<br /> et bestemt job eller tilbud? </b>
	</td>
	</tr>
	<tr><td valign=top colspan=2>
	<select name="jobid" id="jobid" style="width:220px;">
	<option value="0">Nej</option>
	<%
	strSQL = "SELECT jobnavn, id, jobnr FROM job WHERE ((jobstatus = 1 AND fakturerbart = 1) OR jobstatus = 3) AND jobknr = "& kundeid &" ORDER BY jobnavn"
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	if cint(intJobid) = oRec("id") then
	jsel = "SELECTED"
	else
	jsel = ""
	end if%>
	<option value="<%=oRec("id")%>" <%=jsel%>><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</option>
	<%
	oRec.movenext
	wend
	oRec.close %>
	</select>
	</td></tr>
	<tr><td colspan=2 align=right><br><br>
        <input id="Submit1" type="submit" value=" Opdater >> " /></td></tr>
	
	</table>
                                    </form>
	
	<!-- filteros09 slut -->
		                </td>
		                </tr>
		                </table>
		                </div>
	
	<!-- table div -->
	</div>
</div>
	
	
   
	<br>
	<br>
	</div>
	<br><br>&nbsp;
	
	
	<%
	
	case else
	
	
	%>
    



         <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
    Filarkivet kan tage lidt tid at loade, da der kan være mange filer der skal vises.<br />
	Forventet loadtid:
	<%
	
	exp_loadtid = 0
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	ca.: <b>3-20 sek.</b><br />
    <br />&nbsp;
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
    
  
    </td></tr>
   
        </table>

	</div>
	

    

    <%
	
	'*** Er det kunde der er logget på ? ***
	if request("kundelogin") = "1" then
	kundelogin = 1
	else
	kundelogin = 0
	end if
	
	
    if len(trim(request("FM_sog"))) <> 0 OR request("fms") = "1" then
        
        if len(trim(request("FM_sog"))) <> 0 then

        sogeKri = request("FM_sog")
        sogTxt = sogeKri

        else

         sogeKri = "fdsfD3fve#"
        sogTxt = ""

        end if

        response.cookies("tsa")("filersog") = sogTxt

    else

        if jobid <> 0 then
        strSQLjob = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& jobid
        oRec.open strSQLjob, oConn, 3
        if not oRec.EOF then

        sogeKri = oRec("jobnavn")
        sogTxt = sogeKri

        end if
        oRec.close

        else

        if request.cookies("tsa")("filersog") <> "" then
        sogeKri = request.cookies("tsa")("filersog")
        sogTxt = sogeKri
        else
        sogeKri = "fdsfD3fve#"
        sogTxt = ""
        end if

        end if
    end if
	%>
	
	<%if kundelogin <> 1 AND nomenu = 0 then
	sindhTop = 82
	%>
		
		

        <%call menu_2014() %>

	<%else
	%>
		
	
	<%
	sindhTop = 60
	end if%>
    
    

    <%response.flush %>
 
	


	<div id="sindhold" style="position:absolute; left:90px; top:<%=sindhTop%>px; visibility:visible; display:; z-index:50;">
	
    <%
    oimg = "icon_cabinet_open.png"
	oleft = 0
	otop = 0
	owdt = 400
	oskrift = "Filarkiv"
	
	'call sideoverskrift_2013(oleft, otop, owdt, oimg, oskrift)
    
     %>

	
	<%
        pTxt = oskrift
	'call sideindholdtop
	if kundelogin <> 1 then 
	
	call filterheader_2013(20,0,960,pTxt)
	
	%>
    <form action="filer.asp?nomenu=<%=nomenu%>&fms=1" method="post">
	<table width=100% cellspacing=0 cellpadding=3 border=0>


        <tr><td><b>Søg i filarkiv:</b><br /><input type="text" name="FM_sog" value="<%=sogTxt %>" style="width:400px; border:2px #6CAE1C solid;" /><br />
        <span style="font-size:9px; color:#999999;">Søg på kunde, job eller filnavn</span>

          <br /><br />  <input type="CHECKBOX" value="1" name="vistommefoldere" <%=vistommefoldereCHK %> /> Vis tomme foldere 
            </td><td><input type="submit" value=" Vis filer >> "></td></tr>
	</table>
    </form>
	
	
	<!-- filter header -->
	</td></tr></table>
	</div>
	<br /><br /><br />
	
	<%if (level <= 2 OR level = 6) then  
            
            'if kundeid <> 0 then%>
			<a href="javascript:popUp('filer.asp?func=oprfo&kundeid=<%=kundeid%>&jobid=<%=jobid%>&nomenu=<%=nomenu%>','600','500','250','120');" target="_self" class=vmenu><img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0">&nbsp;Opret folder >></a>
		     &nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="javascript:popUp('upload.asp?type=job&id=<%=id%>&kundeid=<%=kundeid%>&jobid=<%=jobid%>&nomenu=<%=nomenu%>','600','500','250','120');" target="_self" class=vmenu><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;Upload fil >></a>
			
            <%'else %>
<!--
            <img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0">&nbsp;Opret folder >> (vælg først kunde)
             &nbsp;&nbsp;|&nbsp;&nbsp;
			<img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;Upload fil >> (vælg først kunde)
			-->
            <%'end if %>
            
           
			<%end if%>
	
	<%end if%>
	
	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 950
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	
	
	
	<table cellspacing=0 bgcolor="#ffffff" cellpadding=2 border=0 width=100%>
	<%
	
	
	jobidSQL2 = ""
	
	
    kundeseSQL = ""
	    
   
	
	
        kundeIdSQL = " k.kkundenavn LIKE '%"& sogeKri &"%' OR k.kkundenr = '"& sogeKri &"'"
        jobIdSQL = " OR j2.jobnavn LIKE '%"& sogeKri &"%' OR j2.jobnr = '"& sogeKri &"'"
        filerIdSQL = " OR fi.filnavn LIKE '%"& sogeKri &"%'"
        folderIdSQL = " OR fo.navn LIKE '%"& sogeKri &"%'"
	
	strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
	&" fi.adg_kunde, fi.adg_admin, fi.adg_alle, fo.id AS foid, fo.kundese, "_
	&" fo.jobid AS jobid, filnavn, fi.id AS fiid, COUNT(fi.id) AS antalfiler,"_
	&" fi.dato, "_
	&" j1.jobnr AS j1jobnr, j1.jobnavn AS j1jobnavn, j1.jobans1 AS j1jobans1, j1.jobans2 AS j1jobans2, "_
	&" j2.jobnr As j2jobnr, j2.jobnavn As j2jobnavn, j2.jobans1 AS j2jobans1, j2.jobans2 AS j2jobans2, "_
	&" fo.dato AS fodato, kkundenavn, kkundenr"_
	&" FROM foldere fo "_
	&" LEFT JOIN filer AS fi ON (fi.folderid = fo.id "& jobidSQL2 &") "_
	&" LEFT JOIN job AS j1 ON (j1.id = fo.jobid) "_
	&" LEFT JOIN job AS j2 ON (j2.id = fi.jobid) "_
    &" LEFT JOIN kunder AS k ON (k.kid = fo.kundeid) "_
	&" WHERE ("& kundeIdSQL &" "& jobIdSQL &" "& filerIdSQL &" "& folderIdSQL &") "& kundeseSQL &" GROUP BY foid, fiid ORDER BY "& filorderby
	'"& gamleFilerKri &"
	
	
	'Response.write strSQL
	'Response.flush
	%>
	
	
	<tr bgcolor="#8caae6">
	<td valign=bottom class=alt width="350"><a href="filer.asp?sortby=navn&kundeid=<%=kundeid %>&id=<%=id %>&jobid=<%=jobid %>&nomenu=<%=nomenu %>&func=<%=func %>&FM_sog=<%=sogTxt %>" class=alt><u>Folder / Filer navn</u></a></td>
	<td valign=bottom align=center class=alt><a href="filer.asp?sortby=dato&kundeid=<%=kundeid %>&id=<%=id %>&jobid=<%=jobid %>&nomenu=<%=nomenu %>&func=<%=func %>&FM_sog=<%=sogTxt %>" class=alt><u>Dato</u></a></td>
	<td valign=bottom class=alt><b>Kontakt</b> (kunde)<br /><b>Job</b></td>
	<td valign=bottom class=alt><b>Ekstern adgang<br>til folder?</b></td>
	
	
	<%
	'** Er det kunde der er logget ind ?
	if kundelogin <> 1 then
	%>
	
	<td valign=bottom class=alt align=center><b>Fil rettigheder.<br>Admin.</b></td>
	<td valign=bottom class=alt align=center><b>Fil rettigheder.<br>Alle medarbjedere.</b></td>
	<td valign=bottom class=alt align=center class=red><b>Slet</b></td>
	<%else%>
	<td colspan=5>&nbsp;</td>
	<%end if%>
	</tr>
	<%
	
	
        
        if sogeKri = "fdsfD3fve#" then
        response.end
        end if
	
    'response.write "strSQL: "& strSQL
    'response.Flush

	x = 0
    y = 0
	lastfolderid = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	'******************** Folder *****************************

    if cint(vistommefoldere) = 1 OR oRec("antalfiler") = "1" then

	if lastfolderid <> oRec("foid") then
	
       y = y + 1

               ' select case right(y, 1)
				'case 0,2,4,6,8
				'folderbg = "#eff3ff"
				'case else
				folderbg = "#FFFFFF"
                'end select

		'** Er det kunde eller medarbejder der er logget ind ?
		if kundelogin <> 1 then 'ikke kunde logget ind => Medarbjeder logget ind
			
			
		'** Tjekker rettigheder eller om man er jobanssvarlig ***
		editok = 0
			
			'** På foldere der er tilknyttet et job ***
			if len(oRec("j1jobnavn")) <> 0 then
			
			
				if level = 1 then '** Administrator
				editok = 1
				else
						'*** jobans 
						if cint(session("mid")) = oRec("j1jobans1") OR cint(session("mid")) = oRec("j1jobans2") OR _
						(cint(oRec("j1jobans1")) = 0 AND cint(oRec("j1jobans2")) = 0) then
						editok = 1
						end if
				end if
			
			else
			'** På foldere der IKKE er tilknyttet et job ***
				
				if level <= 3 OR level = 6 then '** Admin eller niveau 1 må redigere FOLDER
				editok = 1
				else
				editok = 0
				end if
			
			end if
		    
            
		
				%>
				<tr bgcolor="<%=folderbg %>" id="trid_<%=oRec("foid")%>" class="trfo">
					
					<td style="border-top:1px #cccccc solid; height:30px;">
                    
                   
                    <%if oRec("antalfiler") = "1" then 
                    spcls = "showfolder"
                    spcol = "#000000"%>
					<span id="foid_<%=oRec("foid")%>" class="showfolder"><img src="../ill/folder.png"  alt="" border="0"/></span>
                    <%else 
                    spcls = ""
                    spcol = "#999999"%>
                    <img src="../ill/folder_blue.png"  alt="" border="0"/>
					<%end if %>
                    
                  

					
					<span id="fxid_<%=oRec("foid")%>" class="<%=spcls %>" style="font-size:12px; color:<%=spcol %>;"><b><u>+ <%=left(oRec("foldernavn"), 40)%></u></b></span>
					<%if oRec("foid") <> 14 AND oRec("foid") <> 500 AND oRec("foid") <> 1000 AND editok = 1 then%>
                    &nbsp;<a href="javascript:popUp('filer.asp?func=fored&kundeid=<%=oRec("kundeid")%>&id=<%=oRec("foid")%>&nomenu=<%=nomenu%>','600','500','250','120');" target="_self"><img src="../ill/blyant.gif" width="12" height="11" alt="" border="0"></a>
				    <%end if%>

                    
					
					
					
					</td>
					
					<td class=lille style="border-top:1px #cccccc solid; white-space:nowrap;" align=center><!--<%=oRec("fodato")%>-->&nbsp;</td>
					
					<td class=graa valign="top" style="border-top:1px #cccccc solid; padding-top:15px;">

                    <%if len(trim(oRec("kkundenavn"))) <> 0 then%>
					<b><%=left(oRec("kkundenavn"), 30)%> (<%=oRec("kkundenr")%>)</b>
					<%end if%>

					<!--<%if len(trim(oRec("j1jobnavn"))) <> 0 then%>
					<br /><%=left(oRec("j1jobnavn"), 30)%> (<%=oRec("j1jobnr")%>)
					<%end if%>-->
					&nbsp;</td>
					
					<td align=center  style="border-top:1px #cccccc solid;">
					<%if oRec("kundese") = "1" then%>
					<i>V</i>
					<%end if%>&nbsp;
					</td>
					
					<td style="border-top:1px #cccccc solid;" colspan=2>&nbsp;</td>
					<td align=center style="border-top:1px #cccccc solid;">
					<%if oRec("foid") <> 14 AND oRec("foid") <> 500 AND oRec("foid") <> 1000 AND editok = 1 then%>
					<a href="filer.asp?func=slet&kundeid=<%=oRec("kundeid")%>&id=<%=oRec("foid")%>&jobid=<%=jobid%>&nomenu=<%=nomenu%>"><img src="../ill/slet_16.gif" border="0" /></a>
					<%else%>
					&nbsp;
					<%end if%></td>
					
				</tr>
				<%
				
				
				
		
	
	else%>
		<tr bgcolor="<%=folderbg %>">	
			<td style="border-bottom:1px #8caae6 solid;"><img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0"><b><%=oRec("foldernavn")%></b></td>
			<td class=lille style="border-bottom:1px #8caae6 solid;" align=center><%=oRec("fodato")%>&nbsp;</td>
			
            <td class=graa style="border-bottom:1px #8caae6 solid;">
			<%if len(trim(oRec("j1jobnavn"))) <> 0 then%>
			<%=oRec("j1jobnavn")%> (<%=oRec("j1jobnr")%>)
			<%end if%>
			&nbsp;</td>
			
			<td style="border-bottom:1px #8caae6 solid;">
			<%if oRec("kundese") = "1" then%>
			<i>V</i>
			<%end if%>&nbsp;
			</td>
			
			<td style="border-bottom:1px #8caae6 solid;">&nbsp;</td>
			<td style="border-bottom:1px #8caae6 solid;" colspan=4>&nbsp;</td>
		</tr>
			
	<%end if '*** kundelogget ind
		
	end if 'lastfolder%>
	
	
	<%
	'******************** Filer *****************************
	if isNull(oRec("filnavn")) <> true then
		'** Er det kunde der er logget ind ?
		if kundelogin <> 1 then
		
				'*** Rettigheder ***
				
				'***
				'*** Hvis adgang til alle
				'*** Hvis level = admin 
				'*** Hvis Jobans (folder tilknyttet job) (editok = 1)
				'***
				
                'if (oRec("adg_alle") = 1 OR (level = 1) OR editok = 1) AND len(trim(oRec("filnavn"))) <> 0 then
				
				'select case right(x, 1)
				'case 0,2,4,6,8
				bgtd = "#FFFFFF"
				'case else
				'bgtd = "#D6DFf5"
                '#eff3ff
				'end select

                filnavnTxt = replace(oRec("filnavn"), "_txt", ".txt")
				%>
				
				<tr bgcolor="<%=bgtd %>" class="foid_<%=oRec("foid")%>" style="visibility:visible; display:;">	
					<td style="padding-left:24px;"><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;<a href="../inc/upload/<%=lto%>/<%=filnavnTxt%>" class='vmenu' target="_blank"><%=left(filnavnTxt, 40)%></a> [<%=right(filnavnTxt, 3) %>]&nbsp;
					<%if (level = 1 OR oRec("adg_alle") = 1) then%>
					<a href="javascript:popUp('filer.asp?kundeid=<%=oRec("kundeid")%>&jobid=<%=jobid%>&func=redfil&id=<%=oRec("fiid")%>&nomenu=<%=nomenu%>','600','500','250','120');" target="_self"><img src="../ill/blyant.gif" width="12" height="11" alt="" border="0"></a>
					<%end if%>
					&nbsp;</td>
						
						<td class=lille align=center style="white-space:nowrap;"><%=oRec("dato")%>&nbsp;</td>
						
                        

						<td class=graa >
						<%if len(trim(oRec("j2jobnavn"))) <> 0 then%>
					<%=left(oRec("j2jobnavn"), 30)%> (<%=oRec("j2jobnr")%>)
					<%end if%>
                    &nbsp;
						</td>
						
						<td align=center >
						<%if oRec("adg_kunde") = "1" then%>
						<i>V</i>
						<%end if%>
						&nbsp;
						</td>
						
							
							<td align=center >
							<%if oRec("adg_admin") = "1" then%>
							<i>V</i>
							<%end if%>&nbsp;</td>
							
							<td align=center >
							<%if oRec("adg_alle") = "1" then%>
							<i>V</i>
							<%end if%>&nbsp;</td>
							
							<td align=center >
							<%if (level = 1 OR editok = 1) then%>
							<a href="filer.asp?func=sletfil&kundeid=<%=oRec("kundeid")%>&id=<%=oRec("fiid")%>&jobid=<%=jobid%>&nomenu=<%=nomenu%>" target="_blank"><img src="../ill/slet_16.gif" border="0"/></a>
							<%end if%>&nbsp;</td>
				</tr>
				<%'end if '** Rettigheder
				
				
		else '*** Kundelogin
			if len(oRec("filnavn")) <> 0 then%>
			<tr>	
					
					<td style="padding-left:20px; border-bottom:1px #cccccc solid;"><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;<a href="../inc/upload/<%=lto%>/<%=oRec("filnavn")%>" class='vmenulille' target="_blank"><%=oRec("filnavn")%></a>&nbsp;</td>
					<td align=center class=lille ><%=oRec("dato")%></td>
					<td colspan=8 >&nbsp;</td>
			</tr>
			<%end if
		
		end if
		
	end if
	
    lastfilnavn = oRec("filnavn")
	lastfolderid = oRec("foid") 
	
	x = x + 1

    end if 'tommefiler oRec("antalfiler") = "1"


    response.flush
	oRec.movenext
	wend
	oRec.close 
	
	
	%>
	
	</table>
	</div>
	
	<br><br>
	<%if kundelogin <> 1 then%>
   
	<%else%>
	<a href="Javascript:history.back()"><< Tilbage</a>
	<%end if%>
	<br>
	<br>
	&nbsp;
	</div>
	<br><br>&nbsp;
	<%
	
	
	end select
	
	%>
	
	
	
	<%
end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->


<!--- 
******************  Rettigheds Matrix filrarkiv ***********************************************
***********************************************************************************************
			Folder			| 		Filer			| 	Folder tilknyttet job
							|						|	Hvis jobans   	Ikke jobans
			Se 	Opr 	Red | Se	upload	 red.	|  	Se/Opr/Red 		Se/Opr/Red   
***********************************************************************************************
Admin	:	*	*		*	| *		*		*		|	*	*	*		*	*	*
***********************************************************************************************
N1		:	*	*		*	| *		*		-/*		|	*	* 	*		-	-	-
***********************************************************************************************
N2		:	*	-		-	| *		-		-/*		|	*	-	*		-	-	-			
***********************************************************************************************
Kunde	:	*	-		-	| *		-		-		|	*	- 	-		*	-	-
***********************************************************************************************
--->