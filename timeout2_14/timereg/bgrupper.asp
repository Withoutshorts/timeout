<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/medarb_func.asp"-->


<!--#include file="../inc/regular/topmenu_inc.asp"-->





<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	id = request("id") 
	
	%>
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<script language="javascript">
    function alterall(f) {
        for (i = 0; f.elements[i]; i++) {
            e = f.elements[i];
            if (e.type == "checkbox") {
                e.checked = (e.checked) ? false : true;
            }
        }
    }
</script>
	
	<%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->

	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
        
	
	<%call tableDiv(0,0,800)%>


        <h4>Brugergrupper (rettigheder)</h4>

	<table cellspacing="0" cellpadding="2" border="0" width="100%">
	<tr>
    <%
	select case func
	case "med"
	strSQL = "SELECT id, navn FROM brugergrupper WHERE id=" & id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	gruppeNavn = oRec("Navn")
	end if
	oRec.close
	%><form name="brugergruppe" method="post" action="bgrupper.asp?func=update">
	<td valign="top" colspan="3" height=30>
	Medlemmer i <b><%=gruppeNavn%></b> gruppen:</td>
	<td><!--<div align="right"><a href="bgrupper.asp?func=slet&id=<%=id%>"><img src="../ill/slet_16.gif" alt="Slet" border="0"></a></div>--></td>	
	<tr bgcolor="5582D2" style="padding:5px;">
	<td width="40" colspan="3" class=alt><b>Navn</b></td><td align=right valign=top><input type="checkbox" name="all" align="right" valign="bottom" onclick="alterall(this.form)"/></td>
	</tr>
	<%
	strSQL = "SELECT Mnavn, Mid FROM medarbejdere WHERE brugergruppe = "&id&" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	'tæl brugere
	strCountBrg = "SELECT count(*) AS AntBruger FROM medarbejdere WHERE brugergruppe = "&id&""
	set BrgCount = oConn.execute(strCountBrg)
	StrCountBruger = BrgCount("AntBruger")
	
	strCHKIDS = "0"
	while not oRec.EOF
	strID = oRec("Mid")
	strCHKIDS = strCHKIDS & "." & strID
	%>
	<tr style="padding:5px">
		<td height="20" ><a href="medarb_red.asp?menu=medarb&func=red&id=<%=strID%>"><%=oRec("Mnavn")%> </a></td>
		<td colspan=2>&nbsp;</td>
		<td valign="top" align=right><input type="checkbox" name="CHK<%=strID %>"" /></td>
	</tr>
		<tr>
		<td  colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<%
	oRec.movenext
	wend
	oRec.close
	
	%>
	<tr>
	<td colspan="2" style="padding:5px;">brugere i alt: <%=StrCountBruger%><br /><br></td><td></td><td align="right"></td></tr>
	<tr bgcolor="5582D2">
	<td width="40" colspan="4" style="padding:5px;" class=alt><b>Indstillinger</b></td>
	</tr>
	</table>
	<!--<br />Gruppenavn: <input type="text" name="grpnavn" value="<%=gruppeNavn%>" />-->
	<br />
        Flyt valgte medarbejdere til:
        <select name="brggrp" style="width:200px;">
        <%
        	strSQL2 = "SELECT id, navn FROM brugergrupper"
			oRec.open strSQL2, oConn, 3
			while not oRec.EOF 
			strID = oRec("id")
			strNavn = oRec("navn")
			if cint(id) = cint(strID) then
			strselected = "selected"
			end if
			%>

			<option value="<%=strID%>" <%=strselected%>><%=strNavn%></option>
			<%
			strselected = NULL
			oRec.movenext
	        wend
			oRec.close
			%>
        </select>
        &nbsp;<input type="submit" value="<%=tsa_txt_144 %>" />
        <br />
        <% if StrCountBruger <> "0" then
        strDis = "disabled"
        end if %>
        <input type="hidden" name="curgrp" value="<%=id%>" />
        <input type="hidden" name="CheckID" value="<%=STRCHKIDS%>" />
        <br /><br />
        &nbsp;
        </form>
        </div>
        
	<%
	case "update"
	
	'if request("grpnavn") <> "" then
	GlGruppe = request("curgrp")
	NyGruppe = request("brggrp")
	GrpNavn = request("grpnavn")
	
	'UpdateGruppeSQL = "Update brugergrupper SET Navn = '" & GrpNavn & "' WHERE id = " & GlGruppe
	'oConn.Execute(UpdateGruppeSQL)
	
'flyt medarbejdere til ny brugergruppe
if NOT GlGruppe = NyGruppe then	
    idsToCheck = request("CheckID")
	UpdateMedarbSQL = "UPDATE medarbejdere SET Brugergruppe = '" & NyGruppe & "' WHERE mid in ("	
	arr1 = split(idsToCheck,".")
    for i = 1 to ubound(arr1)
    strMedid = arr1(i)
        if request("CHK" & strMedid) = "on" then
            IDLIST = IDLIST & strMedid & ","
        end if
    next
   if len(IDLIST) > 0 then
    IDLIST = left(IDLIST,cint(len(IDLIST)-1))
    UpdateMedarbSQL = UpdateMedarbSQL & IDLIST & ")" 
    oConn.Execute(UpdateMedarbSQL)
   end if
end if
	'end if
	response.Redirect "bgrupper.asp"
	case "slet"%>
	<div id="sindhold" style="position:static;">
	<h3><img src="../ill/ac0019-24.gif" width="24" height="24" alt="" border="0">&nbsp;Brugergruppe - Slet?</h3>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#ffff99">
	<tr>
	    <td style="border:1px red dashed; padding-right:5; padding-left:5; padding-top:5; padding-bottom:5;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"><br>
		Du er ved at <b>slette</b> en brugergruppe. <br>
		Du vil samtidig slette alle indstillinger for denne brugergruppe.<br>
		Dine indstillinger vil <b>ikke kunne genskabes</b>. <br>
		<br>Er du sikker på at du vil slette denne brugergruppe?<br><br>
		<a href="bgrupper.asp?func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
		<br>
		<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
		<br>
		<br>
	</div>
	<%
	case "sletok"
	'tjek om gruppen er tom
	'tæl brugere
	strCountBrg = "SELECT count(*) AS AntBruger FROM medarbejdere WHERE brugergruppe = "&id&""
	set BrgCount = oConn.execute(strCountBrg)
	StrCountBruger = BrgCount("AntBruger")
	
	if strCountBruger = "0" then
	oConn.execute("DELETE FROM brugergrupper WHERE id = "& id)
	response.Redirect "bgrupper.asp"
	else
	response.Write "det er ikke muligt at slette denne gruppe da der stadig er medarbejdere tilnyttet"
	end if
	
	
	case "opret"
	strNavn = request("FM_grpnavn")
	if strNavn <> "" then
	strNavn = Replace(strNavn,"'","''")
	oConn.execute("INSERT INTO brugergrupper (rettigheder, navn) VALUES (1, '"& strNavn  &"')")
	response.Redirect "bgrupper.asp"
	else
	%>
	<div id="sindhold" style="position:absolute; left:20; top:132;">
	<div id="Div1" style="position:absolute; left:20; top:132;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%"  >
	<tr >
		<td width="8" valign=top rowspan=2></td>
		<td colspan=2 valign="top" ><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2 class="style1"></td>
	</tr>
	<tr >
		<td colspan=2 class=alt><b>Brugergruppe info:</b></td>
	</tr>
	
	
	<tr>
	<form action="bgrupper.asp?func=opret" method="post">
	
	<tr>
		<td  >&nbsp;</td>
		<td valign=top>&nbsp;</td>
		
		<td valign=top>&nbsp;</td>
		<td class="style1"  >&nbsp;</td>
	</tr>
	
	<tr>
		<td  >&nbsp;</td>
		<td><font color=red size=2>*</font> Gruppenavn:</td>
		<td><input type="text" name="FM_grpnavn" value="" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td class="style1"  >&nbsp;</td>
	</tr>
	<tr>
		<td >&nbsp;</td>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="200" height="1" alt="" border="0"><input type="submit" value="<%=tsa_txt_144 %>" /></td>
		<td class="style1"  >&nbsp;</td>
	</tr>
	</form>
	</table>
<br><br>
<br>

<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
</div>
	</div>
	<%
	end if
	case else
	%>
	<td valign="top">
	<table cellspacing="0" cellpadding="0" border="0" width="100%" >
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top></td>
		<td colspan="3" valign="top"></td>
	</tr>
	<tr bgcolor="5582D2" height="30px">
		<td width="70%"><a href="bgrupper.asp?menu=medarb&sort=navn" class=alt><b>Navn</b></a></td>
		<td align="center" width="20%"><a href="bgrupper.asp?menu=medarb&sort=nr" class=alt><b>Rettigheds niveau</b></a></td>
		<td width="">&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn, rettigheder FROM brugergrupper ORDER BY navn"
	else
	strSQL = "SELECT id, navn, rettigheder FROM brugergrupper ORDER BY rettigheder"
	end if
	
	x = 0
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	strSQL2 = "SELECT Mid FROM medarbejdere WHERE brugergruppe = "& oRec("id")
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF 
	x = x + 1
	oRec2.movenext
	wend
	oRec2.close
	Antal = x
	%>
	<tr>
		<td  colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr style="padding:5px;">
		<td valign="top"></td>
		<td><a href="bgrupper.asp?menu=medarb&func=med&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>&nbsp;&nbsp;(<%=Antal%>)<br>
		<%
		SELECT CASE oRec("rettigheder")
		case 1
		varBeks = "Administrator gruppen har adgang til alle områder og alle funktioner."
		case 2
		varBeks = "Denne gruppe har adgang til alle områder (som ovenstående) og alle funktioner i Timeout på lederniveau. <b><font color=darkred>Dog ikke adgang til</font></b> TSA/CRM [Indstillinger], og interne timepriser (løn)"
		case 3
		varBeks = "Denne gruppe har adgang til alle områder (som ovenstående) og alle funktioner i Timeout på lederniveau. <b><font color=darkred>Dog ikke adgang til</font></b> TSA [Statistik] og TSA [Fakturering], og begrænset adgang til TSA [Medarbejdere]."
		case 4
		varBeks = "Har adgang til TSA og CRM delen (TSA [timeregistrering] og TSA [medarbejdere], samt CRM [Kalender], CRM [Aktions historik] og CRM [Firma kontakter]).<br> Adgang <b>kun</b> på bruger niveau, det vil sige indtastning af timer, samt oversigter der vedrører medarbejderen selv."
		case 6
		varBeks = "Har kun adgang til TSA delen (som gruppe TSA + CRM niveau 1)."
		case 7
		varBeks = "Har kun adgang til TSA delen (som gruppe TSA + CRM niveau 2)."
		case 8
		varBeks = "Har adgang til TSA [timeregistrering]."
		end select 
		Response.write varBeks
		%></td>
		<td align="center"><%=oRec("rettigheder")%></td>
		
		<!-- udkommenteret funktion -->
		<td><!--<a href="bgrupper.asp?func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="Slet" border="0"></a>--></td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr >
		<td valign="top"></td>
		<td colspan="4" valign="bottom"></td>
	</tr>	
	</table>
	
	<%end select%>
<br>
<br>
</TD>
</TR>
</TABLE>
</div>

<%end if%><!--#include file="../inc/regular/footer_inc.asp"-->