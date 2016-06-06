<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	%>

    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->


	
    <%

    call menu_2014()

    if len(trim(request("FM_sog"))) <> 0 then
    sogTxt = trim(request("FM_sog"))
    else
    sogTxt = ""
    end if

    ptop = 102
	pleft = 20
	pwdt = 802
	
	call filterheader_2013(ptop,pleft,pwdt,pTxt) 

	oskrift = "Søg i fakturaer" %>
	
    
	 <h4><%=oskrift %></h4>
     <form action="erp_fakturaer_find.asp" method="Post">
	<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#FFFFFF">
    <tr><td style="width:400px;">
    Søg på faktura tekst, materiale- eller aktivitets -linje<br />
    (% wildcard)
    </td>
    <td>
    <input type="text" name="FM_sog" value="<%=sogTxt %>" style="width:250px;" />
    </td>
    <td><input type="submit" value="Søg >>" /></td></tr>

    </table>
	</form>


    <!-- filterHeader -->
    </td></tr></table></div>

    <br /><br />



    <%
    
    tTop = 100
	tLeft = 20
	tWdth = 792
	
	
	call tableDiv(tTop,tLeft,tWdth)
     %>
     <table cellspacing=0 cellpadding=2 border=0 width=100%>
    <tr><td>

    <%

    if len(trim(sogTxt)) <> 0 then

    strSogResTxt = "<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor=""#FFFFFF"" >"

    x = 0
    strSogResTxt = strSogResTxt & "<tr><td><h4>Fakturaer:<br><span style=""font-size:11px;"">Fakturanr</span></h4></td><td align=right>&nbsp;</td><td align=right>&nbsp;</td><td align=right>Pris</td><td style=""width:400px; padding-left:20px;"">Tekst</td></tr>"
    '** Fakturaer ***'
    strSQLf = "SELECT fid, aftaleid, jobid, faknr, beloeb AS pris, jobbesk AS faktxt, fakadr, kkundenavn, jobnavn, jobnr FROM fakturaer AS f "_
    &" LEFT JOIN job ON (id = jobid) "_
    &" LEFT JOIN kunder ON (kid = fakadr) "_
    &" WHERE "_
    &" f.faknr LIKE '"& sogTxt &"%' OR f.jobbesk LIKE '%"& sogTxt &"%' ORDER BY faknr DESC LIMIT 100"

    oRec.open strSQLf, oConn, 3
    while not oRec.EOF 

    select case right(x, 1)
    case 0,2,4,6,8
    bgthis = ""
    case else
    bgthis = "#EfF3FF"
    end select

    thisTxt = replace(lcase(oRec("faktxt")), lcase(sogTxt), "<span style='background-color:#FFFF99;'>"& lcase(sogTxt) & "</span>")

    strSogResTxt = strSogResTxt & "<tr style=""background-color:"& bgthis &";""><td valign=top style=""border-bottom:1px #CCCCCC solid; width:200px; white-space:nowrap;"">"_ 
    &""& oRec("kkundenavn") &"<br><b>"& oRec("jobnavn") &"</b> ("& oRec("jobnr") &")<br>"_
    &"<a href=""erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id="&oRec("fid")&"&FM_jobonoff=&FM_kunde="&oRec("fakadr")&"&FM_job="&oRec("jobid")&"&FM_aftale="&oRec("aftaleid")&"&fromfakhist=1"" class=""vmenu"" target=""_blank"">"& oRec("faknr") &"</a></td>"_
    &"<td align=right style=""border-bottom:1px #CCCCCC solid;"">&nbsp;</td>"_
    &"<td align=right style=""border-bottom:1px #CCCCCC solid;"">&nbsp;</td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("pris")) & "</td>"_
    &"<td style=""border-bottom:1px #CCCCCC solid; padding-left:20px;"">"& thisTxt & "</td></tr>"

    x = x + 1
    oRec.movenext
    wend
    oRec.close

    strSogResTxt = strSogResTxt & "<tr><td colspan=5><br><br>&nbsp;</td></tr>"
    strSogResTxt = strSogResTxt & "<tr><td><h4>Aktiviteter:<br><span style=""font-size:11px;"">Fakturanr</span></h4></td><td align=right>Antal</td><td align=right>Stk. pris</td><td align=right>Pris</td><td style=""padding-left:20px;"">Tekst</td></tr>"
    '*** FD - sum-aktiviteter ***'
    strSQLfd = "SELECT fid, aftaleid, jobid, f.faknr, f.beloeb, f.fakadr, fd.beskrivelse AS faktxt, fd.enhedspris AS stkpris, fd.aktpris AS pris, fd.antal AS antal FROM faktura_det AS fd"_
    &" LEFT JOIN fakturaer AS f ON (f.fid = fd.fakid)"_
    &" WHERE fd.beskrivelse LIKE '%"& sogTxt &"%' ORDER BY faknr DESC LIMIT 100"
    
     oRec.open strSQLfd, oConn, 3
    while not oRec.EOF 

     select case right(x, 1)
    case 0,2,4,6,8
    bgthis = ""
    case else
    bgthis = "#EfF3FF"
    end select

    thisTxt = replace(lcase(oRec("faktxt")), lcase(sogTxt), "<span style='background-color:#FFFF99;'>"& lcase(sogTxt) & "</span>")

    strSogResTxt = strSogResTxt & "<tr style=""background-color:"& bgthis &";""><td valign=top style=""border-bottom:1px #CCCCCC solid;"">"_
    &"<a href=""erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id="&oRec("fid")&"&FM_jobonoff=&FM_kunde="&oRec("fakadr")&"&FM_job="&oRec("jobid")&"&FM_aftale="&oRec("aftaleid")&"&fromfakhist=1"" class=""vmenu"" target=""_blank"">"& oRec("faknr") &"</a></td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("antal")) & "</td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("stkpris")) & "</td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("pris")) & "</td><td style=""border-bottom:1px #CCCCCC solid; padding-left:20px;"">"& thisTxt & "</td></tr>"
    
    x = x + 1
    oRec.movenext
    wend
    oRec.close

    strSogResTxt = strSogResTxt & "<tr><td colspan=5><br><br>&nbsp;</td></tr>"
    strSogResTxt = strSogResTxt & "<tr><td><h4>Materialer:<br><span style=""font-size:11px;"">Fakturanr</span></h4></td><td align=right>Antal</td><td align=right>Stk. pris</td><td align=right>Pris</td><td style=""padding-left:20px;"">Tekst</td></tr>"
    '**** Materialer **'
    strSQLfms = "SELECT f.fid, f.aftaleid, f.jobid, f.faknr, f.beloeb, matnavn AS faktxt, matantal AS antal, matenhedspris as stkpris, matbeloeb AS pris FROM fak_mat_spec AS fms "_
    &" LEFT JOIN fakturaer AS f ON (f.fid = fms.matfakid)"_
    &" WHERE fms.matnavn LIKE '%"& sogTxt &"%' ORDER BY faknr DESC LIMIT 100"

    'Response.write strSQLfms
    'Response.flush

     oRec.open strSQLfms, oConn, 3
    while not oRec.EOF 

     select case right(x, 1)
    case 0,2,4,6,8
    bgthis = ""
    case else
    bgthis = "#EfF3FF"
    end select


    thisTxt = replace(oRec("faktxt"), sogTxt, "<span style='background-color:#FFFF99;'>"& sogTxt & "</span>")
    strSogResTxt = strSogResTxt & "<tr style=""background-color:"& bgthis &";""><td valign=top style=""border-bottom:1px #CCCCCC solid;""><a href=""erp_fak_godkendt_2007.asp?id="&oRec("fid")&"&aftid="&oRec("aftaleid")&"&jobid="&oRec("jobid")&""" class=""vmenu"" target=""_blank"">"& oRec("faknr") &"</a></td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("antal")) & "</td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("stkpris")) & "</td>"_
    &"<td valign=top align=right style=""border-bottom:1px #CCCCCC solid;"">"& formatnumber(oRec("pris")) & "</td><td style=""border-bottom:1px #CCCCCC solid; padding-left:20px;"">"& thisTxt & "</td></tr>"

    x = x + 1
    oRec.movenext
    wend
    oRec.close

    strSogResTxt = strSogResTxt & "</table>"

    end if

    %>


    <div id="sogres">


    <%if len(trim(strSogResTxt)) <> 0 then %>
    <%=strSogResTxt %>
    <%else %>
    Brug linjen ovenfor til at søge i eksisterende fakturaer.
    <%end if %>
    </div>
    <br /><br /><br /><br /><br />&nbsp;
     </td></tr>

    </table>

    </div> <!-- table div -->
	
    <br /><br /><br /><br /><br />&nbsp;
	
	
	

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
